"""
PPT 转视频 - 桌面应用入口
使用 pywebview 创建原生窗口，内嵌 Flask 服务
支持 macOS 和 Windows
"""

import multiprocessing
multiprocessing.freeze_support()  # 必须在最顶部：防止 PyInstaller 打包后在 macOS 上重复启动主进程

import sys
import os
import traceback
import threading
import socket
import time

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


def _log_file():
    root = os.path.join(os.path.expanduser('~'), 'PPTtoVideo_data')
    os.makedirs(root, exist_ok=True)
    return os.path.join(root, 'app.log')


def log_message(msg):
    """写一条普通日志到用户数据目录，便于排查问题"""
    try:
        with open(_log_file(), 'a', encoding='utf-8') as f:
            f.write(time.strftime('[%Y-%m-%d %H:%M:%S] ') + msg + '\n')
    except Exception:
        pass


def log_exception(where):
    """把当前异常写入日志文件（打包后无控制台窗口，只能靠日志排错）"""
    try:
        with open(_log_file(), 'a', encoding='utf-8') as f:
            f.write(time.strftime('[%Y-%m-%d %H:%M:%S] ') + where + '\n')
            f.write(traceback.format_exc() + '\n')
    except Exception:
        pass


def find_free_port():
    """找一个可用的端口"""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(('127.0.0.1', 0))
        return s.getsockname()[1]


def start_flask(flask_app, port):
    """在后台线程中运行 Flask 服务，异常写入日志文件"""
    try:
        import logging
        log = logging.getLogger('werkzeug')
        log.setLevel(logging.ERROR)
        flask_app.run(host='127.0.0.1', port=port, debug=False, use_reloader=False)
    except Exception:
        log_exception('Flask 服务异常退出')


def wait_server_ready(url, timeout=60):
    """等待 Flask 真正能返回 HTTP 响应（最多等待 timeout 秒）"""
    import urllib.request
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            with urllib.request.urlopen(url, timeout=1) as resp:
                if resp.status == 200:
                    return True
        except Exception:
            time.sleep(0.5)
    return False


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

    # 先在主线程完成重模块导入（打包后首次启动较慢，含大量 DLL/模块加载），
    # 避免服务还没就绪窗口就打开导致 127.0.0.1 拒绝连接
    try:
        from app import app as flask_app
    except Exception:
        log_exception('导入应用模块失败')
        raise

    # 启动 Flask（后台线程只负责运行）
    flask_thread = threading.Thread(target=start_flask, args=(flask_app, port), daemon=True)
    flask_thread.start()

    # 等待 Flask 真正能提供 HTTP 服务（最多 60 秒）
    if not wait_server_ready(url, timeout=60):
        log_message(f'Flask 服务 60 秒内未就绪: {url}')

    # 创建原生窗口
    try:
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
    except Exception:
        log_exception('创建桌面窗口失败，尝试用系统浏览器打开')
        try:
            import webbrowser
            webbrowser.open(url)
            # 保持主线程存活，让 Flask 继续为浏览器提供页面
            while flask_thread.is_alive():
                time.sleep(1)
        except Exception:
            log_exception('系统浏览器兜底也失败')

    os._exit(0)


if __name__ == '__main__':
    try:
        main()
    except Exception:
        log_exception('程序启动失败')
        sys.exit(1)
