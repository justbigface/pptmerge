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


def clone_slide(src_slide, dest_prs):
    """Clone *src_slide* into *dest_prs* preserving shapes and media.

    This function takes a *source* slide and makes a deep‑copy of its XML
    into a *new* slide in *dest_prs*.  It also copies the relationships
    (images, charts, media, hyperlinks) so nothing goes missing.  It is
    not 100 % of PowerPoint features, but足以覆盖常见场景（文本框、图片、图表、形状、超链接）。
    """
    # 1) 选一个空白版式。大多数主题都有 index 6 = Blank，否则退回 0。
    blank_layout = dest_prs.slide_layouts[6] if len(dest_prs.slide_layouts) > 6 else dest_prs.slide_layouts[0]
    new_slide = dest_prs.slides.add_slide(blank_layout)

    # 2) 深拷贝 <p:spTree> 下的每个 shape 元素
    for shape in src_slide.shapes:
        new_el = copy.deepcopy(shape.element)
        # _spTree.append 比 insert_element_before 简单且安全
        new_slide.shapes._spTree.append(new_el)

    # 3) 复制关键关系（图片 / 图表 / 媒体 / 超链接）
    for rel in src_slide.part.rels.values():
        if rel.reltype in (RT.IMAGE, RT.CHART, RT.MEDIA, RT.HYPERLINK):
            # 让 python‑pptx 自动分配 rId，避免冲突
            new_slide.part.rels.add_relationship(rel.reltype, rel._target, rel.is_external)



def merge_presentations(streams):
    """Concatenate every slide of each PPTX *stream* into one *Presentation*."""
    merged = Presentation()

    # — 删除模板自动生成的空白首页（如果存在）
    if getattr(merged.slides, "_sldIdLst", None) and len(merged.slides) > 0:
        merged.slides._sldIdLst.remove(merged.slides._sldIdLst[0])

    for s in streams:
        prs = Presentation(s)
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
    streams = []
    try:
        # —— 下载全部 PPTX ——
        for url in urls:
            resp = requests.get(url, stream=True, timeout=30)
            resp.raise_for_status()
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.pptx')
            for chunk in resp.iter_content(8192):
                tmp.write(chunk)
            tmp.close()
            temp_paths.append(tmp.name)

        # —— 打开文件流并合并 ——
        for p in temp_paths:
            streams.append(open(p, 'rb'))

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
        # 关闭并删除临时文件
        for s in streams:
            try:
                s.close()
            except Exception:
                pass
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
