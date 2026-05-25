"""
PPT 转视频工具
功能：将 PPT 每页转为图片，配合文本生成语音，合成视频
"""

import os
import sys
import asyncio
import subprocess
import tempfile
import shutil
from pathlib import Path

import edge_tts
from moviepy import ImageClip, AudioFileClip, concatenate_videoclips
from PIL import Image


# ========== 配置 ==========
# edge-tts 中文语音，模拟新闻主播效果
# 可选: zh-CN-YunxiNeural (男声), zh-CN-XiaoxiaoNeural (女声)
TTS_VOICE = "zh-CN-YunxiNeural"
TTS_RATE = "+0%"  # 语速调整，如 "+10%", "-10%"

# 视频输出配置
OUTPUT_FPS = 24
OUTPUT_CODEC = "libx264"


def ppt_to_images(ppt_path: str, output_dir: str) -> list[str]:
    """
    将 PPT 文件转换为图片（每页一张）
    需要系统安装 LibreOffice
    """
    ppt_path = os.path.abspath(ppt_path)
    output_dir = os.path.abspath(output_dir)
    os.makedirs(output_dir, exist_ok=True)

    # 使用 LibreOffice 将 PPT 转为 PDF，再转为图片
    # 先转 PDF
    pdf_dir = tempfile.mkdtemp()
    try:
        subprocess.run(
            [
                "soffice",
                "--headless",
                "--convert-to", "pdf",
                "--outdir", pdf_dir,
                ppt_path,
            ],
            check=True,
            capture_output=True,
            text=True,
        )
    except FileNotFoundError:
        print("错误：未找到 LibreOffice (soffice)，请先安装 LibreOffice")
        print("macOS 安装：brew install --cask libreoffice")
        sys.exit(1)

    # 找到生成的 PDF 文件
    pdf_files = list(Path(pdf_dir).glob("*.pdf"))
    if not pdf_files:
        print("错误：PPT 转 PDF 失败")
        sys.exit(1)

    pdf_path = str(pdf_files[0])

    # 使用 pdf2image 或 soffice 直接导出图片
    # 这里用另一种方式：直接用 soffice 导出为 PNG
    shutil.rmtree(pdf_dir)

    # 直接将 PPT 转为图片（使用 LibreOffice）
    img_temp_dir = tempfile.mkdtemp()
    subprocess.run(
        [
            "soffice",
            "--headless",
            "--convert-to", "png",
            "--outdir", img_temp_dir,
            ppt_path,
        ],
        capture_output=True,
        text=True,
    )

    # LibreOffice 只能将整个 PPT 转为一个图片，我们需要用另一种方式
    # 更好的方法：先转 PDF，再用 Python 将 PDF 每页转为图片
    shutil.rmtree(img_temp_dir)

    # 重新转 PDF
    pdf_dir = tempfile.mkdtemp()
    subprocess.run(
        [
            "soffice",
            "--headless",
            "--convert-to", "pdf",
            "--outdir", pdf_dir,
            ppt_path,
        ],
        check=True,
        capture_output=True,
        text=True,
    )
    pdf_path = str(list(Path(pdf_dir).glob("*.pdf"))[0])

    # 使用 pdf2image 将 PDF 转为图片
    try:
        from pdf2image import convert_from_path
        images = convert_from_path(pdf_path, dpi=200)
    except ImportError:
        print("错误：需要安装 pdf2image 和 poppler")
        print("安装：pip install pdf2image")
        print("macOS 还需要：brew install poppler")
        sys.exit(1)

    image_paths = []
    for i, img in enumerate(images):
        img_path = os.path.join(output_dir, f"slide_{i + 1:03d}.png")
        img.save(img_path, "PNG")
        image_paths.append(img_path)
        print(f"  已导出第 {i + 1} 页")

    shutil.rmtree(pdf_dir)
    return image_paths


async def text_to_speech(text: str, output_path: str, voice: str = TTS_VOICE, rate: str = TTS_RATE):
    """
    使用 edge-tts 将文本转为语音
    """
    communicate = edge_tts.Communicate(text, voice, rate=rate)
    await communicate.save(output_path)


async def generate_all_audio(texts: list[str], output_dir: str) -> list[str]:
    """
    批量将文本列表转为语音文件
    """
    os.makedirs(output_dir, exist_ok=True)
    audio_paths = []

    for i, text in enumerate(texts):
        if not text.strip():
            # 空文本跳过，但保留占位
            audio_paths.append(None)
            print(f"  第 {i + 1} 页：无文本，将使用静默")
            continue

        audio_path = os.path.join(output_dir, f"audio_{i + 1:03d}.mp3")
        await text_to_speech(text.strip(), audio_path)
        audio_paths.append(audio_path)
        print(f"  第 {i + 1} 页：语音生成完成")

    return audio_paths


def create_video(image_paths: list[str], audio_paths: list[str], output_path: str, default_duration: float = 3.0):
    """
    将图片和音频合成为视频
    - 每张图片显示时长 = 对应音频时长
    - 如果没有音频，使用默认时长
    """
    clips = []

    for i, (img_path, audio_path) in enumerate(zip(image_paths, audio_paths)):
        if audio_path and os.path.exists(audio_path):
            # 有音频：图片时长 = 音频时长
            audio_clip = AudioFileClip(audio_path)
            duration = audio_clip.duration
            img_clip = ImageClip(img_path).with_duration(duration).with_audio(audio_clip)
        else:
            # 无音频：使用默认时长
            img_clip = ImageClip(img_path).with_duration(default_duration)

        clips.append(img_clip)
        print(f"  第 {i + 1} 页：时长 {img_clip.duration:.1f}s")

    # 拼接所有片段
    print("\n正在合成最终视频...")
    final_video = concatenate_videoclips(clips, method="compose")
    final_video.write_videofile(
        output_path,
        fps=OUTPUT_FPS,
        codec=OUTPUT_CODEC,
        audio_codec="aac",
        logger="bar",
    )
    print(f"\n视频已生成：{output_path}")
    print(f"总时长：{final_video.duration:.1f}s")


def parse_texts(text_content: str) -> list[str]:
    """
    解析文本内容，按换行分割为每页的文本
    空行会被保留为空文本（对应页面使用静默）
    """
    lines = text_content.split("\n")
    # 去掉最后的空行
    while lines and not lines[-1].strip():
        lines.pop()
    return lines


def main():
    import argparse

    parser = argparse.ArgumentParser(description="PPT 转视频工具")
    parser.add_argument("ppt", help="PPT 文件路径 (.pptx 或 .ppt)")
    parser.add_argument("text", help="文本文件路径（每行对应一页 PPT 的旁白）")
    parser.add_argument("-o", "--output", default="output.mp4", help="输出视频路径 (默认: output.mp4)")
    parser.add_argument("-v", "--voice", default=TTS_VOICE, help=f"TTS 语音 (默认: {TTS_VOICE})")
    parser.add_argument("-r", "--rate", default=TTS_RATE, help=f"TTS 语速 (默认: {TTS_RATE})")
    parser.add_argument("-d", "--default-duration", type=float, default=3.0,
                        help="无音频时默认显示时长(秒) (默认: 3.0)")

    args = parser.parse_args()

    # 验证输入文件
    if not os.path.exists(args.ppt):
        print(f"错误：PPT 文件不存在：{args.ppt}")
        sys.exit(1)
    if not os.path.exists(args.text):
        print(f"错误：文本文件不存在：{args.text}")
        sys.exit(1)

    # 创建工作目录
    work_dir = tempfile.mkdtemp(prefix="ppt_video_")
    images_dir = os.path.join(work_dir, "images")
    audio_dir = os.path.join(work_dir, "audio")

    try:
        # 1. PPT 转图片
        print("=" * 50)
        print("步骤 1/3：PPT 转图片")
        print("=" * 50)
        image_paths = ppt_to_images(args.ppt, images_dir)
        print(f"共导出 {len(image_paths)} 页\n")

        # 2. 文本转语音
        print("=" * 50)
        print("步骤 2/3：文本转语音")
        print("=" * 50)
        with open(args.text, "r", encoding="utf-8") as f:
            text_content = f.read()
        texts = parse_texts(text_content)

        if len(texts) != len(image_paths):
            print(f"警告：文本行数 ({len(texts)}) 与 PPT 页数 ({len(image_paths)}) 不一致")
            # 自动补齐或截断
            if len(texts) < len(image_paths):
                texts.extend([""] * (len(image_paths) - len(texts)))
            else:
                texts = texts[:len(image_paths)]

        audio_paths = asyncio.run(generate_all_audio(texts, audio_dir, ))
        print()

        # 3. 合成视频
        print("=" * 50)
        print("步骤 3/3：合成视频")
        print("=" * 50)
        # 确保音频列表长度与图片一致
        while len(audio_paths) < len(image_paths):
            audio_paths.append(None)

        create_video(image_paths, audio_paths, args.output, args.default_duration)

    finally:
        # 清理临时文件
        shutil.rmtree(work_dir, ignore_errors=True)

    print("\n完成！")


if __name__ == "__main__":
    main()
