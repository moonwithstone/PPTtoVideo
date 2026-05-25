"""
PPT 转视频 - 桌面应用入口
使用 pywebview 创建原生窗口，内嵌 Flask 服务
支持 macOS 和 Windows
"""

import multiprocessing
multiprocessing.freeze_support()  # 必须在最顶部：防止 PyInstaller 打包后在 macOS 上重复启动主进程

import sys
import os
import shutil
import threading
import socket

# 确保打包后模块路径正确
if getattr(sys, 'frozen', False):
    os.environ['PATH'] = os.path.dirname(sys.executable) + os.pathsep + os.environ.get('PATH', '')
    # 将进程工作目录切到可写的用户数据目录，避免任何库以相对路径写入到 .app 包内的只读位置
    _data_root = os.path.join(os.path.expanduser('~'), 'PPTtoVideo_data')
    os.makedirs(_data_root, exist_ok=True)
    try:
        os.chdir(_data_root)
    except Exception:
        pass


def find_free_port():
    """找一个可用的端口"""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(('127.0.0.1', 0))
        return s.getsockname()[1]


def start_flask(port):
    """在后台线程中启动 Flask 服务"""
    from app import app
    import logging
    log = logging.getLogger('werkzeug')
    log.setLevel(logging.ERROR)
    app.run(host='127.0.0.1', port=port, debug=False, use_reloader=False)


class Api:
    """暴露给前端的原生 API"""

    def __init__(self, window):
        self.window = window

    def pick_save_path(self):
        """弹出原生保存对话框，让用户选择视频保存位置"""
        import webview

        save_path = self.window.create_file_dialog(
            webview.SAVE_DIALOG,
            directory=os.path.expanduser('~/Desktop'),
            save_filename='output.mp4',
            file_types=('视频文件 (*.mp4)',)
        )

        if save_path:
            if isinstance(save_path, (list, tuple)):
                save_path = save_path[0]
            if not save_path.endswith('.mp4'):
                save_path += '.mp4'
            return {"success": True, "path": save_path}

        return {"success": False}


def main():
    port = find_free_port()
    url = f'http://127.0.0.1:{port}'

    # 启动 Flask（后台线程）
    flask_thread = threading.Thread(target=start_flask, args=(port,), daemon=True)
    flask_thread.start()

    # 等待 Flask 启动
    import time
    for _ in range(50):
        try:
            with socket.create_connection(('127.0.0.1', port), timeout=0.1):
                break
        except (ConnectionRefusedError, OSError):
            time.sleep(0.1)

    # 创建原生窗口
    import webview

    window = webview.create_window(
        title='PPT 转视频',
        url=url,
        width=900,
        height=750,
        min_size=(700, 600),
        confirm_close=False,
        text_select=True,
    )

    # 绑定原生 API
    api = Api(window)
    window.expose(api.pick_save_path)

    # 启动 webview
    webview.start()

    os._exit(0)


if __name__ == '__main__':
    main()
