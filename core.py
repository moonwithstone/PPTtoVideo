"""
核心功能模块：PPT转图片、文本转语音、合成视频
自动检测已安装的办公软件（WPS / MS Office / LibreOffice）
使用 PyMuPDF 替代 poppler，无需额外系统依赖
"""

import os
import sys
import asyncio
import subprocess
import tempfile
import shutil
import platform
from pathlib import Path

import edge_tts
import fitz  # PyMuPDF
from moviepy import ImageClip, AudioFileClip, concatenate_videoclips


# ========== 默认配置 ==========
TTS_VOICE = "zh-CN-YunxiNeural"
TTS_RATE = "+0%"
OUTPUT_FPS = 24
OUTPUT_CODEC = "libx264"

IS_WINDOWS = platform.system() == 'Windows'
IS_MAC = platform.system() == 'Darwin'


def _find_soffice() -> str | None:
    """查找 soffice 可执行文件路径（兼容 macOS app bundle 未加入 PATH 的情况）"""
    if shutil.which("soffice"):
        return "soffice"
    if IS_MAC:
        candidates = [
            "/Applications/LibreOffice.app/Contents/MacOS/soffice",
            os.path.expanduser("~/Applications/LibreOffice.app/Contents/MacOS/soffice"),
        ]
        for p in candidates:
            if os.path.exists(p):
                return p
    if IS_WINDOWS:
        candidates = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ]
        for p in candidates:
            if os.path.exists(p):
                return p
    return None


# ==================================================
#  办公软件检测与 PPT 转换
# ==================================================

def _find_mac_app(app_name):
    """macOS: 检测应用是否安装"""
    try:
        result = subprocess.run(
            ['mdfind', f'kMDItemKind == "Application" && kMDItemDisplayName == "{app_name}"'],
            capture_output=True, text=True, timeout=5
        )
        return bool(result.stdout.strip())
    except Exception:
        return False


def _find_mac_app_path(app_name):
    """macOS: 查找应用路径"""
    common_paths = [
        f'/Applications/{app_name}.app',
        f'/Applications/Office/{app_name}.app',
        os.path.expanduser(f'~/Applications/{app_name}.app'),
    ]
    for p in common_paths:
        if os.path.exists(p):
            return p
    # 用 mdfind 搜索
    try:
        result = subprocess.run(
            ['mdfind', f'kMDItemFSName == "{app_name}.app" && kMDItemKind == "Application"'],
            capture_output=True, text=True, timeout=5
        )
        paths = result.stdout.strip().split('\n')
        if paths and paths[0]:
            return paths[0]
    except Exception:
        pass
    return None


def detect_office_software():
    """
    检测系统中安装的办公软件
    返回: list[dict] 按优先级排序，每项包含 name, type, available
    """
    results = []

    if IS_WINDOWS:
        # Windows: 通过 COM 检测
        try:
            import comtypes.client
            # 检测 WPS
            try:
                comtypes.client.CreateObject("KWPP.Application")
                results.append({"name": "WPS Office", "type": "wps_com", "available": True})
            except Exception:
                pass
            # 检测 MS Office
            try:
                comtypes.client.CreateObject("PowerPoint.Application")
                results.append({"name": "Microsoft Office", "type": "ms_com", "available": True})
            except Exception:
                pass
        except ImportError:
            pass

        # 检测 LibreOffice (Windows)
        lo_paths = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ]
        for lo in lo_paths:
            if os.path.exists(lo):
                results.append({"name": "LibreOffice", "type": "libreoffice", "available": True, "path": lo})
                break

    elif IS_MAC:
        # macOS: 检测各办公软件
        # LibreOffice 优先（唯一无界面的方式，不会弹出 app 窗口）
        soffice = _find_soffice()
        if soffice:
            results.append({"name": "LibreOffice", "type": "libreoffice", "available": True})

        # WPS
        wps_path = _find_mac_app_path("WPS Office")
        if wps_path:
            results.append({"name": "WPS Office", "type": "wps_mac", "available": True, "path": wps_path})

        # Microsoft PowerPoint
        ppt_path = _find_mac_app_path("Microsoft PowerPoint")
        if ppt_path:
            results.append({"name": "Microsoft PowerPoint", "type": "ms_mac", "available": True, "path": ppt_path})

        # Keynote (macOS 自带，但对 pptx 兼容性一般)
        if os.path.exists("/Applications/Keynote.app"):
            results.append({"name": "Keynote", "type": "keynote", "available": True})

    else:
        # Linux
        if _find_soffice():
            results.append({"name": "LibreOffice", "type": "libreoffice", "available": True})

    return results


def check_dependencies():
    """检查系统依赖是否满足"""
    issues = []
    offices = detect_office_software()
    if not offices:
        if IS_WINDOWS:
            issues.append("未检测到办公软件（WPS / MS Office / LibreOffice），请至少安装其中一个")
        elif IS_MAC:
            issues.append("未检测到办公软件（WPS / MS Office / Keynote / LibreOffice），请至少安装其中一个")
        else:
            issues.append("未检测到 LibreOffice，请安装：sudo apt install libreoffice")
    return issues


def _convert_ppt_windows_com(ppt_path: str, output_dir: str, com_type: str, progress_cb=None) -> list[str]:
    """Windows: 使用 COM 自动化将 PPT 另存为 PDF，再用 PyMuPDF 转图片
    比逐页 Export 更稳定，WPS 和 MS Office 均支持 SaveAs PDF（格式常量 32）
    """
    import comtypes
    import comtypes.client

    def log(msg):
        if progress_cb:
            progress_cb(msg)
        print(msg)

    ppt_path = os.path.abspath(ppt_path)
    software_name = "WPS Office" if com_type == "wps_com" else "Microsoft Office"
    app_name = "KWPP.Application" if com_type == "wps_com" else "PowerPoint.Application"
    log(f"使用 {software_name} 转换 PPT...")

    # COM 必须在当前线程初始化（Flask 后台线程不会自动初始化）
    comtypes.CoInitialize()
    pdf_dir = tempfile.mkdtemp()
    pdf_path = os.path.join(pdf_dir, Path(ppt_path).stem + ".pdf")

    try:
        app = comtypes.client.CreateObject(app_name)
        app.Visible = False
        try:
            presentation = app.Presentations.Open(ppt_path, WithWindow=False)
            try:
                # 32 = ppSaveAsPDF，WPS 和 MS Office 通用常量
                presentation.SaveAs(pdf_path, 32)
                log(f"{software_name} 导出 PDF 成功")
            finally:
                presentation.Close()
        finally:
            try:
                app.Quit()
            except Exception:
                pass
    finally:
        comtypes.CoUninitialize()

    if not os.path.exists(pdf_path):
        shutil.rmtree(pdf_dir, ignore_errors=True)
        raise RuntimeError(f"{software_name} 导出 PDF 失败，请确认文件未被其他程序占用")

    image_paths = pdf_to_images(pdf_path, output_dir, progress_cb)
    shutil.rmtree(pdf_dir, ignore_errors=True)
    return image_paths


def _convert_ppt_mac_applescript(ppt_path: str, output_dir: str, app_type: str, progress_cb=None) -> list[str]:
    """macOS: 使用 AppleScript 调用 PowerPoint/Keynote 导出"""
    def log(msg):
        if progress_cb:
            progress_cb(msg)
        print(msg)

    ppt_path = os.path.abspath(ppt_path)

    if app_type == "ms_mac":
        # Microsoft PowerPoint for Mac
        log("使用 Microsoft PowerPoint 转换...")
        script = f'''
        tell application "Microsoft PowerPoint"
            open "{ppt_path}"
            set thePresentation to active presentation
            set slideCount to count of slides of thePresentation
            repeat with i from 1 to slideCount
                set theSlide to slide i of thePresentation
                set exportPath to "{output_dir}/slide_" & text -3 thru -1 of ("000" & i) & ".png"
                save theSlide in exportPath as save as PNG
            end repeat
            close thePresentation saving no
        end tell
        '''
    elif app_type == "keynote":
        # Keynote - 导出整个演示文稿为图片
        # 需要处理打开 .pptx 时的兼容性警告弹窗
        # 使用 try/error + front document 提高兼容性（适配 Keynote 变种如 Keynote Creator Studio）
        log("使用 Keynote 转换...")
        script = f'''
        tell application "Keynote"
            activate
            open POSIX file "{ppt_path}"
        end tell

        -- 多等一会儿，让文件加载和弹窗出现
        delay 5

        -- 尝试关闭可能出现的警告/兼容性弹窗（多次尝试）
        tell application "System Events"
            tell process "Keynote"
                repeat 5 times
                    try
                        if exists sheet 1 of window 1 then
                            click button 1 of sheet 1 of window 1
                        end if
                    end try
                    delay 1
                end repeat
            end tell
        end tell

        delay 2

        -- 使用 try/error 重试获取文档（避免使用 count of documents）
        set theDoc to missing value
        repeat 20 times
            try
                tell application "Keynote"
                    set theDoc to front document
                end tell
                exit repeat
            on error
                delay 1
            end try
        end repeat

        if theDoc is missing value then
            tell application "Keynote" to quit
            error "Keynote 无法打开文档（可能是格式不兼容）"
        end if

        tell application "Keynote"
            try
                export theDoc as slide images to POSIX file "{output_dir}" with properties {{image format:PNG, skipped slides:false}}
            on error errMsg
                try
                    close theDoc saving no
                end try
                quit
                error errMsg
            end try
            try
                close theDoc saving no
            end try
            quit
        end tell
        '''
    else:
        raise RuntimeError(f"不支持的 macOS 应用类型: {app_type}")

    try:
        subprocess.run(['osascript', '-e', script], check=True, capture_output=True, text=True, timeout=180)
    except subprocess.CalledProcessError as e:
        raise RuntimeError(f"AppleScript 执行失败: {e.stderr}")
    except subprocess.TimeoutExpired:
        raise RuntimeError("转换超时（超过 3 分钟）")

    # 收集生成的图片
    image_paths = sorted([
        os.path.join(output_dir, f)
        for f in os.listdir(output_dir)
        if f.lower().endswith('.png')
    ])

    for i, p in enumerate(image_paths):
        log(f"已导出第 {i + 1}/{len(image_paths)} 页图片")

    return image_paths


def _convert_ppt_libreoffice(ppt_path: str, output_dir: str, progress_cb=None) -> list[str]:
    """使用 LibreOffice 转 PDF，再用 PyMuPDF 转图片"""
    def log(msg):
        if progress_cb:
            progress_cb(msg)
        print(msg)

    ppt_path = os.path.abspath(ppt_path)
    log("使用 LibreOffice 转换...")

    # PPT → PDF
    soffice = _find_soffice()
    if not soffice:
        raise RuntimeError("未找到 LibreOffice，请安装后重试（macOS 请从 libreoffice.org 下载 .dmg 安装包）")

    pdf_dir = tempfile.mkdtemp()
    try:
        subprocess.run(
            [
                soffice,
                "--headless",       # 无界面模式
                "--norestore",      # 不恢复上次崩溃的会话窗口
                "--nologo",         # 不显示启动 Logo
                "--nofirststartwizard",  # 跳过首次启动向导
                "--convert-to", "pdf",
                "--outdir", pdf_dir,
                ppt_path,
            ],
            check=True, capture_output=True, text=True, timeout=120,
        )
    except FileNotFoundError:
        shutil.rmtree(pdf_dir, ignore_errors=True)
        raise RuntimeError("未找到 LibreOffice (soffice)")

    pdf_files = list(Path(pdf_dir).glob("*.pdf"))
    if not pdf_files:
        shutil.rmtree(pdf_dir, ignore_errors=True)
        raise RuntimeError("PPT 转 PDF 失败")

    # PDF → 图片（PyMuPDF，纯 Python，无需 poppler）
    image_paths = pdf_to_images(str(pdf_files[0]), output_dir, progress_cb)
    shutil.rmtree(pdf_dir, ignore_errors=True)
    return image_paths


def _convert_ppt_wps_mac(ppt_path: str, output_dir: str, progress_cb=None) -> list[str]:
    """macOS: 使用 WPS 命令行工具转 PDF，再用 PyMuPDF 转图片"""
    def log(msg):
        if progress_cb:
            progress_cb(msg)
        print(msg)

    ppt_path = os.path.abspath(ppt_path)
    log("使用 WPS Office 转换...")

    pdf_dir = tempfile.mkdtemp()
    pdf_path = os.path.join(pdf_dir, Path(ppt_path).stem + ".pdf")

    # WPS 可能的命令行工具路径（不同版本位置不同）
    wps_cli_candidates = [
        "/Applications/WPS Office.app/Contents/MacOS/wpp",           # 新版演示模块
        "/Applications/WPS Office.app/Contents/MacOS/wpspdf",        # 旧版专用工具
        "/Applications/WPS Office.app/Contents/MacOS/wps",           # 通用入口
        "/Applications/WPS Office.app/Contents/SharedSupport/wpp",
        "/Applications/WPS Office.app/Contents/SharedSupport/wpspdf",
        os.path.expanduser("~/Applications/WPS Office.app/Contents/MacOS/wpp"),
        os.path.expanduser("~/Applications/WPS Office.app/Contents/MacOS/wpspdf"),
    ]

    for wps_cli in wps_cli_candidates:
        if not os.path.exists(wps_cli):
            continue
        # 先尝试 --headless 模式（部分版本支持，与 LibreOffice 用法一致）
        try:
            result = subprocess.run(
                [wps_cli, "--headless", "--convert-to", "pdf", "--outdir", pdf_dir, ppt_path],
                capture_output=True, timeout=120
            )
            if result.returncode == 0:
                pdf_files = list(Path(pdf_dir).glob("*.pdf"))
                if pdf_files:
                    log("WPS 命令行（headless 模式）转换成功")
                    image_paths = pdf_to_images(str(pdf_files[0]), output_dir, progress_cb)
                    shutil.rmtree(pdf_dir, ignore_errors=True)
                    return image_paths
        except Exception:
            pass
        # 再尝试直接传入输出路径的格式
        try:
            result = subprocess.run(
                [wps_cli, ppt_path, pdf_path],
                capture_output=True, timeout=120
            )
            if result.returncode == 0 and os.path.exists(pdf_path):
                log("WPS 命令行转换成功")
                image_paths = pdf_to_images(pdf_path, output_dir, progress_cb)
                shutil.rmtree(pdf_dir, ignore_errors=True)
                return image_paths
        except Exception:
            pass

    shutil.rmtree(pdf_dir, ignore_errors=True)
    raise RuntimeError(
        "WPS macOS 版命令行转换失败。\n"
        "解决方法：在 WPS 中将 PPT 另存为 PDF，再上传 PDF 文件（本工具支持直接上传 PDF）"
    )


def pdf_to_images(pdf_path: str, output_dir: str, progress_cb=None) -> list[str]:
    """使用 PyMuPDF 将 PDF 转为图片（纯 Python，不需要 poppler）"""
    def log(msg):
        if progress_cb:
            progress_cb(msg)
        print(msg)

    doc = fitz.open(pdf_path)
    image_paths = []

    for i in range(len(doc)):
        page = doc[i]
        # 高质量渲染：2x 缩放
        mat = fitz.Matrix(2.0, 2.0)
        pix = page.get_pixmap(matrix=mat)
        img_path = os.path.join(output_dir, f"slide_{i + 1:03d}.png")
        pix.save(img_path)
        image_paths.append(img_path)
        log(f"已导出第 {i + 1}/{len(doc)} 页图片")

    doc.close()
    return image_paths


def ppt_to_images(ppt_path: str, output_dir: str, progress_cb=None) -> list[str]:
    """
    智能文件转图片：
    - PDF 文件：直接用 PyMuPDF 转换（无需任何办公软件）
    - PPT 文件：自动检测并使用已安装的办公软件
    """
    os.makedirs(output_dir, exist_ok=True)

    def log(msg):
        if progress_cb:
            progress_cb(msg)
        print(msg)

    ext = Path(ppt_path).suffix.lower()

    # PDF 文件：直接转图片，不需要任何办公软件
    if ext == '.pdf':
        log("检测到 PDF 文件，直接转换为图片...")
        return pdf_to_images(ppt_path, output_dir, progress_cb)

    # PPT/PPTX 文件：需要办公软件
    offices = detect_office_software()
    if not offices:
        raise RuntimeError(
            "未检测到可用的办公软件，请尝试以下方案：\n"
            "1. 将 PPT 导出为 PDF 后再上传（推荐，最稳定）\n"
            "2. 安装 WPS Office / Microsoft Office / LibreOffice"
        )

    # 按优先级尝试
    last_error = None
    for office in offices:
        try:
            log(f"检测到 {office['name']}，正在使用...")
            otype = office['type']

            if otype == "wps_com":
                return _convert_ppt_windows_com(ppt_path, output_dir, "wps_com", progress_cb)
            elif otype == "ms_com":
                return _convert_ppt_windows_com(ppt_path, output_dir, "ms_com", progress_cb)
            elif otype == "ms_mac":
                return _convert_ppt_mac_applescript(ppt_path, output_dir, "ms_mac", progress_cb)
            elif otype == "keynote":
                return _convert_ppt_mac_applescript(ppt_path, output_dir, "keynote", progress_cb)
            elif otype == "wps_mac":
                return _convert_ppt_wps_mac(ppt_path, output_dir, progress_cb)
            elif otype == "libreoffice":
                return _convert_ppt_libreoffice(ppt_path, output_dir, progress_cb)

        except Exception as e:
            last_error = e
            log(f"{office['name']} 转换失败: {e}，尝试下一个...")
            continue

    raise RuntimeError(f"所有办公软件转换均失败，最后一次错误: {last_error}")


# ==================================================
#  文本转语音
# ==================================================

async def text_to_speech(text: str, output_path: str, voice: str = TTS_VOICE, rate: str = TTS_RATE):
    """使用 edge-tts 将文本转为语音"""
    communicate = edge_tts.Communicate(text, voice, rate=rate)
    await communicate.save(output_path)


async def generate_all_audio(texts: list[str], output_dir: str, voice: str = TTS_VOICE,
                              rate: str = TTS_RATE, progress_cb=None) -> list[str]:
    """批量将文本列表转为语音文件，失败时自动重试，最终失败则用静默代替"""
    os.makedirs(output_dir, exist_ok=True)
    audio_paths = []

    def log(msg):
        if progress_cb:
            progress_cb(msg)
        print(msg)

    for i, text in enumerate(texts):
        if not text.strip():
            audio_paths.append(None)
            log(f"第 {i + 1}/{len(texts)} 页：无文本，将使用静默")
            continue

        audio_path = os.path.join(output_dir, f"audio_{i + 1:03d}.mp3")
        success = False
        last_err = None

        for attempt in range(3):
            try:
                await text_to_speech(text.strip(), audio_path, voice, rate)
                success = True
                break
            except Exception as e:
                last_err = e
                if attempt < 2:
                    log(f"第 {i + 1}/{len(texts)} 页：语音生成失败，正在重试（{attempt + 1}/3）...")
                    await asyncio.sleep(1.5)

        if success:
            audio_paths.append(audio_path)
            log(f"第 {i + 1}/{len(texts)} 页：语音生成完成")
        else:
            audio_paths.append(None)
            log(f"第 {i + 1}/{len(texts)} 页：语音生成失败，将使用静默（错误：{last_err}）")

    return audio_paths


# ==================================================
#  视频合成
# ==================================================

def create_video(image_paths: list[str], audio_paths: list[str], output_path: str,
                 default_duration: float = 3.0, intro_delay: float = 0.0, progress_cb=None):
    """将图片和音频合成为视频
    intro_delay: 第一页画面出现后，延迟多少秒再开始播放音频（默认 0 = 无延迟）
    """
    clips = []

    def log(msg):
        if progress_cb:
            progress_cb(msg)
        print(msg)

    for i, (img_path, audio_path) in enumerate(zip(image_paths, audio_paths)):
        if audio_path and os.path.exists(audio_path):
            audio_clip = AudioFileClip(audio_path)
            duration = audio_clip.duration
            img_clip = ImageClip(img_path).with_duration(duration).with_audio(audio_clip)
        else:
            img_clip = ImageClip(img_path).with_duration(default_duration)

        # 第一页：在正式片段前插入一段静音片段
        if i == 0 and intro_delay > 0:
            silent_clip = ImageClip(img_path).with_duration(intro_delay)
            clips.append(silent_clip)
            log(f"第 1 页：插入 {intro_delay:.1f}s 静音前导")

        clips.append(img_clip)
        log(f"第 {i + 1}/{len(image_paths)} 页：时长 {img_clip.duration:.1f}s")

    log("正在合成最终视频...")
    final_video = concatenate_videoclips(clips, method="compose")

    # 关键修复：先让 MoviePy 全程在可写的临时目录下工作，完成后再 move 到用户选的路径
    # 避免输出路径同目录只读、包含特殊字符、或为相对路径时，MoviePy 生成 outputTEMP_MPY_*.mp4 临时文件写入失败
    output_path = os.path.abspath(output_path)
    work_dir = tempfile.mkdtemp(prefix='ppt2video_render_')
    temp_output = os.path.join(work_dir, 'output.mp4')
    temp_audio = os.path.join(work_dir, 'temp_audio.m4a')

    # 切换到可写工作目录，防止 ffmpeg 以相对路径写临时文件到当前 cwd（.app 包内是只读的）
    original_cwd = os.getcwd()
    try:
        os.chdir(work_dir)

        final_video.write_videofile(
            temp_output,
            fps=OUTPUT_FPS,
            codec=OUTPUT_CODEC,
            audio_codec="aac",
            temp_audiofile=temp_audio,
            remove_temp=True,
            logger="bar",
        )

        # 移动到用户指定的最终位置
        log(f"保存视频到: {output_path}")
        os.makedirs(os.path.dirname(output_path) or '.', exist_ok=True)
        if os.path.exists(output_path):
            os.remove(output_path)
        shutil.move(temp_output, output_path)
    finally:
        os.chdir(original_cwd)
        shutil.rmtree(work_dir, ignore_errors=True)

    total_duration = final_video.duration
    log(f"视频生成完成！总时长：{total_duration:.1f}s")
    return total_duration


def parse_texts(text_content: str) -> list[str]:
    """解析文本内容：按空行分页，同一页内多行自动拼接为一段连续旁白

    规则：
    - 一个或多个连续空行 = 页面分隔符
    - 同一页内的多行会用空格拼接为一段文本
    - 首尾空行忽略
    """
    import re
    # 以一个或多个空行作为分隔符（空行 = 仅含空白字符的行）
    blocks = re.split(r'\n\s*\n', text_content.strip())
    pages = []
    for block in blocks:
        # 同一页内的多行去除首尾空白，用空格拼接
        merged = ' '.join(line.strip() for line in block.split('\n') if line.strip())
        if merged:
            pages.append(merged)
    return pages


def generate_video_full(ppt_path: str, text_content: str, output_path: str,
                         voice: str = TTS_VOICE, rate: str = TTS_RATE,
                         default_duration: float = 3.0, intro_delay: float = 0.0,
                         progress_cb=None):
    """
    完整流程：PPT + 文本 → 视频
    返回: (成功: bool, 消息: str)
    """
    def log(msg):
        if progress_cb:
            progress_cb(msg)
        print(msg)

    work_dir = tempfile.mkdtemp(prefix="ppt_video_")
    images_dir = os.path.join(work_dir, "images")
    audio_dir = os.path.join(work_dir, "audio")

    try:
        # 1. PPT 转图片
        log("【步骤 1/3】PPT 转图片...")
        image_paths = ppt_to_images(ppt_path, images_dir, progress_cb)
        log(f"共导出 {len(image_paths)} 页图片")

        # 2. 文本转语音
        log("【步骤 2/3】文本转语音...")
        texts = parse_texts(text_content)

        if len(texts) != len(image_paths):
            log(f"⚠ 文本行数 ({len(texts)}) 与 PPT 页数 ({len(image_paths)}) 不一致，已自动调整")
            if len(texts) < len(image_paths):
                texts.extend([""] * (len(image_paths) - len(texts)))
            else:
                texts = texts[:len(image_paths)]

        audio_paths = asyncio.run(generate_all_audio(texts, audio_dir, voice, rate, progress_cb))

        # 3. 合成视频
        log("【步骤 3/3】合成视频...")
        while len(audio_paths) < len(image_paths):
            audio_paths.append(None)

        total_duration = create_video(image_paths, audio_paths, output_path, default_duration, intro_delay, progress_cb)
        return True, f"视频生成成功！总时长 {total_duration:.1f} 秒"

    except Exception as e:
        return False, f"生成失败：{str(e)}"
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)
