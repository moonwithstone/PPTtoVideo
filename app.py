"""
PPT 转视频 - Web 应用
Flask 后端：文件上传、视频生成、下载
支持独立运行和打包为桌面应用
"""

import os
import sys
import uuid
import threading
import tempfile
from datetime import datetime

from flask import Flask, render_template, request, jsonify, send_file
from werkzeug.utils import secure_filename

from core import generate_video_full, check_dependencies, detect_office_software


def get_base_dir():
    """获取基础目录（兼容 PyInstaller 打包）"""
    if getattr(sys, 'frozen', False):
        return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))


def get_data_dir():
    """获取数据目录（存放上传和输出文件）"""
    if getattr(sys, 'frozen', False):
        # 打包后使用用户目录下的专用文件夹
        data_dir = os.path.join(os.path.expanduser('~'), 'PPTtoVideo_data')
    else:
        data_dir = os.path.dirname(os.path.abspath(__file__))
    os.makedirs(data_dir, exist_ok=True)
    return data_dir


BASE_DIR = get_base_dir()
DATA_DIR = get_data_dir()

app = Flask(__name__,
            template_folder=os.path.join(BASE_DIR, 'templates'),
            static_folder=os.path.join(BASE_DIR, 'static'))
app.config['MAX_CONTENT_LENGTH'] = 200 * 1024 * 1024  # 最大 200MB
app.config['UPLOAD_FOLDER'] = os.path.join(DATA_DIR, 'uploads')
app.config['OUTPUT_FOLDER'] = os.path.join(DATA_DIR, 'outputs')

# 确保目录存在
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

# 任务状态存储
tasks = {}


class Task:
    def __init__(self, task_id):
        self.id = task_id
        self.status = "pending"  # pending, processing, completed, failed
        self.progress = []
        self.result_path = None
        self.error = None
        self.created_at = datetime.now()

    def add_progress(self, message):
        self.progress.append(message)

    def to_dict(self):
        return {
            "id": self.id,
            "status": self.status,
            "progress": self.progress,
            "error": self.error,
            "has_result": self.result_path is not None,
        }


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/check', methods=['GET'])
def check_env():
    """检查系统依赖和办公软件"""
    issues = check_dependencies()
    offices = detect_office_software()
    return jsonify({
        "ok": len(issues) == 0,
        "issues": issues,
        "offices": [o["name"] for o in offices],
    })


@app.route('/generate', methods=['POST'])
def generate():
    """生成视频"""
    # 获取文件（PPT 或 PDF）
    if 'ppt_file' not in request.files:
        return jsonify({"error": "请上传 PPT 或 PDF 文件"}), 400

    ppt_file = request.files['ppt_file']
    if ppt_file.filename == '':
        return jsonify({"error": "请选择文件"}), 400

    # 获取文本内容
    text_content = ""
    if 'text_file' in request.files and request.files['text_file'].filename:
        text_file = request.files['text_file']
        text_content = text_file.read().decode('utf-8')
    elif request.form.get('text_content'):
        text_content = request.form['text_content']
    else:
        return jsonify({"error": "请提供旁白文本（上传文件或在编辑区输入）"}), 400

    # 获取选项
    voice = request.form.get('voice', 'zh-CN-YunxiNeural')
    rate = request.form.get('rate', '+0%')
    intro_delay = float(request.form.get('intro_delay', '0'))
    save_path = request.form.get('save_path', '')  # 用户指定的保存路径

    # 确定输出路径
    task_id = str(uuid.uuid4())[:8]
    if save_path:
        output_path = save_path
    else:
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], f"{task_id}_output.mp4")

    # 保存上传的文件
    task_dir = os.path.join(app.config['UPLOAD_FOLDER'], task_id)
    os.makedirs(task_dir, exist_ok=True)
    # 保留原始扩展名（secure_filename 会把中文名全部去掉，连扊号也丢）
    original_ext = os.path.splitext(ppt_file.filename)[1].lower()
    if original_ext not in ('.pdf', '.ppt', '.pptx'):
        original_ext = '.pdf'  # 默认当 PDF 处理
    ppt_path = os.path.join(task_dir, f"upload{original_ext}")
    ppt_file.save(ppt_path)

    # 创建任务
    task = Task(task_id)
    tasks[task_id] = task

    # 在后台线程中执行
    def run_task():
        task.status = "processing"
        task.add_progress("任务开始...")

        success, message = generate_video_full(
            ppt_path=ppt_path,
            text_content=text_content,
            output_path=output_path,
            voice=voice,
            rate=rate,
            intro_delay=intro_delay,
            progress_cb=task.add_progress,
        )

        if success:
            task.status = "completed"
            task.result_path = output_path
        else:
            task.status = "failed"
            task.error = message

    thread = threading.Thread(target=run_task, daemon=True)
    thread.start()

    return jsonify({"task_id": task_id})


@app.route('/task/<task_id>', methods=['GET'])
def task_status(task_id):
    """查询任务状态"""
    task = tasks.get(task_id)
    if not task:
        return jsonify({"error": "任务不存在"}), 404
    return jsonify(task.to_dict())


@app.route('/download/<task_id>', methods=['GET'])
def download(task_id):
    """下载生成的视频"""
    task = tasks.get(task_id)
    if not task or not task.result_path:
        return jsonify({"error": "视频不存在"}), 404

    if not os.path.exists(task.result_path):
        return jsonify({"error": "视频文件已被清理"}), 404

    return send_file(task.result_path, as_attachment=True,
                     download_name=os.path.basename(task.result_path).split('_', 1)[-1])


@app.route('/voices', methods=['GET'])
def list_voices():
    """返回可用的中文语音列表"""
    voices = [
        {"id": "zh-CN-YunxiNeural",   "name": "云希（男声·阳光）", "gender": "male"},
        {"id": "zh-CN-YunjianNeural", "name": "云健（男声·激情）", "gender": "male"},
        {"id": "zh-CN-YunyangNeural", "name": "云扬（男声·专业）", "gender": "male"},
        {"id": "zh-CN-YunxiaNeural",  "name": "云夏（男声·可爱）", "gender": "male"},
        {"id": "zh-CN-XiaoxiaoNeural","name": "晓晓（女声·温柔）", "gender": "female"},
        {"id": "zh-CN-XiaoyiNeural",  "name": "晓伊（女声·活力）", "gender": "female"},
    ]
    return jsonify(voices)


if __name__ == '__main__':
    print("=" * 50)
    print("  PPT 转视频工具 - Web 版")
    print("  访问 http://127.0.0.1:5000")
    print("=" * 50)
    app.run(debug=True, host='127.0.0.1', port=5000)
