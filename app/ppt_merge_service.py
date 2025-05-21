from flask import Flask, request, send_file, jsonify
import io
import os
import tempfile
import requests
import copy
import contextlib
import urllib.parse
import time
import concurrent.futures
from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT

app = Flask(__name__)

# --- 安全配置 ------------------------------------------------------------------
ALLOWED_HOSTS = {h.lower().strip(): None for h in os.getenv("ALLOWED_HOSTS", "justbigface.fun").split(",") if h}
ALLOWED_PORTS = {int(p) for p in os.getenv("ALLOWED_PORTS", "80,443,9000,9001").split(",") if p}
MAX_SIZE_MB = int(os.getenv("MAX_SIZE_MB", "30"))
MAX_SIZE_BYTES = MAX_SIZE_MB * 1024 * 1024

# --- 工具函数 ------------------------------------------------------------------
def _validate_url(url: str) -> None:
    """如不符合白名单规则则抛出 ValueError"""
    parsed = urllib.parse.urlparse(url)
    host = (parsed.hostname or "").lower()
    port = parsed.port or (443 if parsed.scheme == "https" else 80)
    if host not in ALLOWED_HOSTS:
        raise ValueError(f"Host not allowed: {host}")
    if port not in ALLOWED_PORTS:
        raise ValueError(f"Port not allowed: {port} for URL {url}")

def clone_slide(src_slide, dest_prs):
    """Clone *src_slide* into *dest_prs* preserving shapes and media."""
    blank_layout = dest_prs.slide_layouts[6] if len(dest_prs.slide_layouts) > 6 else dest_prs.slide_layouts[0]
    new_slide = dest_prs.slides.add_slide(blank_layout)
    for shape in src_slide.shapes:
        new_el = copy.deepcopy(shape.element)
        new_slide.shapes._spTree.append(new_el)
    for rId, rel in src_slide.part.rels.items():
        try:
            new_slide.part.rels.add_relationship(rel.reltype, rel._target, is_external=rel.is_external)
        except Exception:
            new_slide.part.rels.add_relationship(rel.reltype, rel._target, is_external=rel.is_external)

def merge_presentations(streams):
    merged = Presentation()
    if len(merged.slides) == 1 and len(merged.slides[0].shapes._spTree) <= 1:
        merged.slides._sldIdLst.remove(merged.slides._sldIdLst[0])
    for s in streams:
        prs = Presentation(s)
        for slide in prs.slides:
            clone_slide(slide, merged)
    return merged

@app.route('/merge_pptx', methods=['POST'])
def merge_pptx():
    start_time = time.time()
    data = request.get_json(silent=True) or {}
    urls = data.get('urls', [])
    app.logger.info(f"Received merge request for URLs: {urls}")
    if not urls:
        app.logger.warning("No URLs provided in merge request.")
        return jsonify({'error': 'No urls provided'}), 400
    if len(urls) < 2:
        app.logger.warning("Less than two URLs provided in merge request.")
        return jsonify({'error': 'Need at least two URLs'}), 400
    temp_paths = []
    try:
        with contextlib.ExitStack() as stack:
            streams = []
            with requests.Session() as session:
                def download_file(url):
                    try:
                        _validate_url(url)
                    except ValueError as ve:
                        app.logger.warning(str(ve))
                        return {'error': str(ve)}
                    try:
                        resp = session.get(url, stream=True, timeout=(5, 30))
                        resp.raise_for_status()
                        content_type = resp.headers.get('Content-Type', '')
                        if 'application/vnd.openxmlformats-officedocument.presentationml.presentation' not in content_type and not url.lower().endswith('.pptx'):
                            return {'error': 'Invalid Content-Type or file extension'}
                        tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.pptx')
                        downloaded_size = 0
                        for chunk in resp.iter_content(8192):
                            downloaded_size += len(chunk)
                            if downloaded_size > MAX_SIZE_BYTES:
                                tmp.close()
                                os.remove(tmp.name)
                                return {'error': f'File size exceeds limit ({MAX_SIZE_MB} MB)'}
                            tmp.write(chunk)
                        tmp.close()
                        return {'success': tmp.name}
                    except Exception as e:
                        app.logger.exception(f"Error downloading {url}")
                        return {'error': 'Error downloading file'}
                max_workers = min(5, len(urls))
                with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
                    future_to_url = {executor.submit(download_file, url): url for url in urls}
                    for future in concurrent.futures.as_completed(future_to_url):
                        url = future_to_url[future]
                        try:
                            result = future.result()
                        except Exception as exc:
                            app.logger.exception(f"Exception in downloading {url}")
                            for p in temp_paths:
                                try:
                                    os.remove(p)
                                except OSError:
                                    pass
                            return jsonify({'error': 'Failed to download file'}), 400
                        if 'success' in result:
                            temp_paths.append(result['success'])
                        else:
                            for p in temp_paths:
                                try:
                                    os.remove(p)
                                except OSError:
                                    pass
                            return jsonify({'error': result.get('error', 'Failed to download file')}), 400
            for p in temp_paths:
                streams.append(stack.enter_context(open(p, 'rb')))
            merged_prs = merge_presentations(streams)
            out = io.BytesIO()
            merged_prs.save(out)
            out.seek(0)
            end_time = time.time()
            duration = end_time - start_time
            app.logger.info(f"Successfully merged {len(urls)} files into {len(merged_prs.slides)} slides in {duration:.2f} seconds.")
            return send_file(
                out,
                as_attachment=True,
                download_name='merged.pptx',
                mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
            )
    except Exception as exc:
        app.logger.error(f"Error merging PPTX files from {urls}: {exc}", exc_info=True)
        return jsonify({'error': 'An internal error occurred during merging.'}), 500
    finally:
        for p in temp_paths:
            try:
                os.remove(p)
            except OSError:
                app.logger.warning(f"Could not remove temporary file: {p}", exc_info=True)

def index():
    return 'PPT Merge Service is running!'

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080)
