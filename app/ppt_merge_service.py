from flask import Flask, request, send_file, jsonify
import io
import os
import tempfile
import requests
from pptx import Presentation
from werkzeug.utils import secure_filename

app = Flask(__name__)

def copy_slide(source_presentation, slide, target_presentation):
    # 复制幻灯片：保持布局一致，但更复杂内容的复制需要深入shape分析
    slide_layout = target_presentation.slide_layouts[0]
    new_slide = target_presentation.slides.add_slide(slide_layout)
    for shape in slide.shapes:
        if shape.shape_type == 1:  # Placeholder or TextBox
            txBox = new_slide.shapes.add_textbox(shape.left, shape.top, shape.width, shape.height)
            txBox.text = shape.text
        elif shape.shape_type == 13:  # Picture
            image_stream = shape.image.blob
            new_slide.shapes.add_picture(io.BytesIO(image_stream), shape.left, shape.top, shape.width, shape.height)
        # 如果需要支持更多shape类型，请扩展此处

def merge_presentations(streams):
    merged = Presentation()
    # 清除默认第一页空白
    if len(merged.slides) > 0:
        sldIdLst = merged.slides._sldIdLst
        sldIdLst.remove(sldIdLst[0])
    for stream in streams:
        pres = Presentation(stream)
        for slide in pres.slides:
            copy_slide(pres, slide, merged)
    return merged

@app.route('/merge_pptx', methods=['POST'])
def merge_pptx():
    data = request.get_json()
    urls = data.get('urls', [])
    if not urls:
        return jsonify({'error': 'No urls provided'}), 400

    # 下载所有pptx到临时文件
    temp_files = []
    for url in urls:
        try:
            r = requests.get(url, stream=True, timeout=10)
            r.raise_for_status()
            tmpf = tempfile.NamedTemporaryFile(delete=False, suffix='.pptx')
            for chunk in r.iter_content(chunk_size=8192):
                tmpf.write(chunk)
            tmpf.close()
            temp_files.append(tmpf.name)
        except Exception as e:
            # 清理已下载的临时文件
            for path in temp_files:
                try:
                    os.remove(path)
                except:
                    pass
            return jsonify({'error': f'Failed downloading {url}: {str(e)}'}), 500

    # 读取临时文件为流并合并
    input_streams = []
    for path in temp_files:
        with open(path, 'rb') as f:
            input_streams.append(io.BytesIO(f.read()))

    merged_ppt = merge_presentations(input_streams)

    # 清理临时文件
    for path in temp_files:
        try:
            os.remove(path)
        except:
            pass

    output_stream = io.BytesIO()
    merged_ppt.save(output_stream)
    output_stream.seek(0)
    return send_file(
        output_stream,
        as_attachment=True,
        download_name='merged.pptx',
        mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
    )

@app.route('/')
def hello():
    return "PPT Merge Service is running!"

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080 )