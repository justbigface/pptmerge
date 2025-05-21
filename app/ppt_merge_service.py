from flask import Flask, request, send_file, jsonify
import io
import os
import tempfile
import requests
import copy
import urllib.parse
from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT

app = Flask(__name__)

# --- 安全配置 ------------------------------------------------------------------
ALLOWED_HOSTS = {h.lower().strip(): None for h in os.getenv("ALLOWED_HOSTS", "justbigface.fun").split(",") if h}
ALLOWED_PORTS = {int(p) for p in os.getenv("ALLOWED_PORTS", "80,443,9000,9001").split(",") if p}
MAX_SIZE_MB = int(os.getenv("MAX_SIZE_MB", "30"))
MAX_SIZE_BYTES = MAX_SIZE_MB * 1024 * 1024

def _validate_url(url: str) -> None:
    parsed = urllib.parse.urlparse(url)
    host = (parsed.hostname or "").lower()
    port = parsed.port or (443 if parsed.scheme == "https" else 80)
    if host not in ALLOWED_HOSTS:
        raise ValueError(f"Host not allowed: {host}")
    if port not in ALLOWED_PORTS:
        raise ValueError(f"Port not allowed: {port} for URL {url}")

def _blank_slide(prs: Presentation):
    idx = 6 if len(prs.slide_layouts) > 6 else 0
    slide = prs.slides.add_slide(prs.slide_layouts[idx])
    for shape in list(slide.shapes):
        slide.shapes._spTree.remove(shape.element)
    return slide

def _clone_slide(src_slide, dst_prs):
    dst_slide = _blank_slide(dst_prs)
    for shp in src_slide.shapes:
        dst_slide.shapes._spTree.append(copy.deepcopy(shp.element))
    for rel in src_slide.part.rels.values():
        if rel.reltype not in (RT.IMAGE, RT.CHART, RT.MEDIA, RT.HYPERLINK):
            continue
        rels = dst_slide.part.rels
        try:
            rels.add_relationship(rel.reltype, rel._target, rId=None, external=rel.is_external)
        except (TypeError, AttributeError):
            rels.add_rel(rel.reltype, rel._target, rel.is_external)

def _merge_presentations(streams):
    base_prs = Presentation(streams[0])
    for stream in streams[1:]:
        src_prs = Presentation(stream)
        for slide in src_prs.slides:
            _clone_slide(slide, base_prs)
    return base_prs

@app.route('/merge_pptx', methods=['POST'])
def merge_pptx():
    data = request.get_json(silent=True) or {}
    urls = data.get('urls') or []
    if not urls:
        return jsonify({'error': 'No urls provided'}), 400
    temp_paths, streams = [], []
    session = requests.Session()
    try:
        for url in urls:
            _validate_url(url)
            resp = session.get(url, stream=True, timeout=(5, 30))
            resp.raise_for_status()
            size = 0
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.pptx')
            for chunk in resp.iter_content(65536):
                size += len(chunk)
                if size > MAX_SIZE_MB * 1024 * 1024:
                    raise ValueError(f'File too large (> {MAX_SIZE_MB} MB): {url}')
                tmp.write(chunk)
            tmp.close()
            temp_paths.append(tmp.name)
        for path in temp_paths:
            streams.append(open(path, 'rb'))
        merged = _merge_presentations(streams)
        buf = io.BytesIO()
        merged.save(buf)
        buf.seek(0)
        return send_file(buf, as_attachment=True, download_name='merged.pptx',
                         mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation')
    except Exception as exc:
        app.logger.exception('merge_pptx failed')
        return jsonify({'error': str(exc)}), 500
    finally:
        for s in streams:
            try:
                s.close()
            except Exception:
                pass
        for p in temp_paths:
            try:
                os.remove(p)
            except Exception:
                pass

@app.route('/')
def index():
    return 'PPT Merge Service is running!'

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080)
