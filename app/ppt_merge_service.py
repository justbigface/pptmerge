from flask import Flask, request, send_file, jsonify
import io
import os
import tempfile
import requests
import copy
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.opc.constants import RELATIONSHIP_TYPE as RT

app = Flask(__name__)


def clone_slide(source_slide, target_prs):
    """Deep‑clone *source_slide* into *target_prs*, preserving shapes, text,
    images, charts and hyperlinks.
    """
    # choose a blank layout in the target presentation
    blank_layout = target_prs.slide_layouts[6] if len(target_prs.slide_layouts) > 6 else target_prs.slide_layouts[0]
    new_slide = target_prs.slides.add_slide(blank_layout)

    # ---clone the XML of every shape---
    for shape in source_slide.shapes:
        new_el = copy.deepcopy(shape.element)
        new_slide.shapes._spTree.insert_element_before(new_el, 'p:extLst')

    # ---clone relationships so pictures / charts don’t break---
    for rel in source_slide.part.rels:
        if rel.reltype in (RT.IMAGE, RT.CHART, RT.MEDIA, RT.HYPERLINK):
            new_slide.part.rels.add_rel(rel.reltype, rel._target, rel.rId)


def merge_presentations(streams):
    """Return a *Presentation* obtained by concatenating all slides of the
    PPTX *streams* (iterable of BytesIO)."""
    merged = Presentation()

    # remove the default blank slide the template comes with
    merged.slides._sldIdLst.remove(merged.slides._sldIdLst[0])

    for stream in streams:
        prs = Presentation(stream)
        for slide in prs.slides:
            clone_slide(slide, merged)
    return merged


@app.route('/merge_pptx', methods=['POST'])
def merge_pptx():
    data = request.get_json(silent=True) or {}
    urls = data.get('urls', [])
    if not urls:
        return jsonify({'error': 'No urls provided'}), 400

    temp_paths = []
    try:
        # ---download all PPTX files---
        for url in urls:
            resp = requests.get(url, stream=True, timeout=20)
            resp.raise_for_status()
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.pptx')
            for chunk in resp.iter_content(8192):
                tmp.write(chunk)
            tmp.close()
            temp_paths.append(tmp.name)

        # ---open streams & merge---
        streams = [open(p, 'rb') for p in temp_paths]
        merged_prs = merge_presentations(streams)

        out = io.BytesIO()
        merged_prs.save(out)
        out.seek(0)
        return send_file(
            out,
            as_attachment=True,
            download_name='merged.pptx',
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
        )
    except Exception as exc:
        return jsonify({'error': str(exc)}), 500
    finally:
        for p in temp_paths:
            try:
                os.remove(p)
            except OSError:
                pass


@app.route('/')
def index():
    return 'PPT Merge Service is running!'


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080)
