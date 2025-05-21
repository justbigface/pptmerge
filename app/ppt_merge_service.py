from flask import Flask, request, send_file, jsonify
import io
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

@app.route('/merge', methods=['POST'])
def merge_ppts():
    files = request.files.getlist('files')
    if not files or len(files) < 2:
        return jsonify({'error': '请上传至少两个pptx文件，且参数名为files'}), 400

    input_streams = []
    for file in files:
        filename = secure_filename(file.filename)
        if not filename.lower().endswith('.pptx'):
            return jsonify({'error': f'文件 {filename} 不是pptx格式'}), 400
        input_streams.append(io.BytesIO(file.read()))

    merged_ppt = merge_presentations(input_streams)
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
    app.run(host='0.0.0.0', port=5001 )