import argparse
import json
import re
import shutil
import subprocess
import time
import zipfile
from pathlib import Path

import pyautogui
import pyperclip
import pygetwindow as gw
import argostranslate.translate
from lxml import etree


# ============================================================
# 固定路径
# ============================================================

PDF24_EXE = r"C:\Users\926\Desktop\PDF24\pdf24-Toolbox.exe"

DEFAULT_INPUT = r"C:\Users\926\Desktop\T5A_Chapter 8.1 Building_TC4"
DEFAULT_OUTPUT = r"C:\Users\926\Desktop\trans"

WD_ROOT = r"C:\Users\926\Desktop\wd"

DELAY_BETWEEN_FILES_SECONDS = 2 * 60
MAX_FAILED_FILES = 10


# ============================================================
# PDF24 界面坐标
# ============================================================

HOME_CONVERT_TO_PDF_X = 1300
HOME_CONVERT_TO_PDF_Y = 250

HOME_PDF_TO_OTHER_X = 1580
HOME_PDF_TO_OTHER_Y = 250

SELECT_FILE_X = 900
SELECT_FILE_Y = 275

HOME_BUTTON_X = 65
HOME_BUTTON_Y = 50


# ============================================================
# PDF → Word 页面坐标
# ============================================================

FORMAT_DROPDOWN_X = 172
FORMAT_DROPDOWN_Y = 527

WORD_DOCX_OPTION_X = 187
WORD_DOCX_OPTION_Y = 910

PDF_TO_WORD_ORANGE_CONVERT_X = 1672
PDF_TO_WORD_ORANGE_CONVERT_Y = 529

PDF_TO_WORD_SAVE_BUTTON_X = 844
PDF_TO_WORD_SAVE_BUTTON_Y = 487


# ============================================================
# Word → PDF 页面坐标
# ============================================================

WORD_TO_PDF_ORANGE_CONVERT_X = 964
WORD_TO_PDF_ORANGE_CONVERT_Y = 545

WORD_TO_PDF_SAVE_BUTTON_X = 825
WORD_TO_PDF_SAVE_BUTTON_Y = 477


# ============================================================
# Word XML 翻译设置
# ============================================================

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}

WORD_XML_PREFIXES = (
    "word/document.xml",
    "word/header",
    "word/footer",
    "word/footnotes.xml",
    "word/endnotes.xml",
    "word/comments.xml",
)


# ============================================================
# 基础函数
# ============================================================

def wait(seconds=1):
    time.sleep(seconds)


def paste_text(text):
    pyperclip.copy(str(text))
    pyautogui.hotkey("ctrl", "v")
    wait(0.3)


def open_pdf24():
    print("正在打开 PDF24 Toolbox...")
    subprocess.Popen([PDF24_EXE])
    wait(5)

    pyautogui.hotkey("alt", "space")
    wait(0.2)
    pyautogui.press("x")
    wait(1)


def close_pdf24():
    print("正在关闭 PDF24...")

    try:
        pyautogui.hotkey("alt", "f4")
        wait(2)
    except Exception:
        pass

    try:
        subprocess.run(
            ["taskkill", "/F", "/IM", "pdf24-Toolbox.exe"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    except Exception:
        pass

    wait(2)


def go_home():
    pyautogui.click(HOME_BUTTON_X, HOME_BUTTON_Y)
    wait(2)


def windows_open_file_dialog(file_path: Path):
    wait(1)
    paste_text(str(file_path))
    pyautogui.press("enter")
    wait(3)


def active_window_title():
    try:
        win = gw.getActiveWindow()
        if win:
            return win.title or ""
    except Exception:
        pass
    return ""


def is_save_dialog_active():
    title = active_window_title()
    lower_title = title.lower()

    keywords = [
        "另存为",
        "保存",
        "save as",
        "save",
    ]

    return any(k.lower() in lower_title for k in keywords)


def wait_for_save_dialog(timeout: int = 30):
    for _ in range(timeout * 2):
        title = active_window_title()
        if title:
            print(f"当前活动窗口：{title}")

        if is_save_dialog_active():
            return True

        time.sleep(0.5)

    return False


def force_click_pdf24_save_button(x: int, y: int, description: str):
    """
    专门处理 PDF24 蓝色保存按钮：
    鼠标指到了但普通 click 不触发时，用多点 mouseDown/mouseUp。
    """
    print(f"正在强制点击：{description}")

    pyautogui.click(x, y)
    wait(0.6)

    points = [
        (x, y),
        (x + 8, y),
        (x - 8, y),
        (x, y + 8),
        (x + 8, y + 8),
        (x - 8, y + 8),
    ]

    for idx, (px, py) in enumerate(points, start=1):
        print(f"  第 {idx} 次尝试点击保存按钮：({px}, {py})")

        pyautogui.moveTo(px, py, duration=0.25)
        wait(0.4)

        pyautogui.mouseDown(button="left")
        wait(0.18)
        pyautogui.mouseUp(button="left")

        wait(1.2)

        if is_save_dialog_active():
            print("已检测到另存为窗口。")
            return True

    if wait_for_save_dialog(timeout=5):
        print("已检测到另存为窗口。")
        return True

    return False


def windows_save_as_dialog(output_path: Path):
    """
    Windows 另存为窗口：
    检测到窗口后，Alt+N 进入文件名框，粘贴完整路径保存。
    """
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    print("等待 Windows 另存为窗口出现...")

    if not wait_for_save_dialog(timeout=30):
        raise RuntimeError("没有检测到 Windows 另存为窗口，停止保存，避免误操作 Ctrl+A")

    wait(1)

    pyautogui.hotkey("alt", "n")
    wait(0.5)

    pyautogui.hotkey("ctrl", "a")
    wait(0.2)

    paste_text(str(output_path))
    wait(0.5)

    pyautogui.press("enter")
    wait(2)

    pyautogui.press("y")
    wait(6)


def wait_for_file_exists(file_path: Path, timeout: int = 60):
    file_path = Path(file_path)

    for _ in range(timeout):
        if file_path.exists() and file_path.stat().st_size > 0:
            return True
        time.sleep(1)

    return False


def wait_between_files(index: int, total: int):
    if index < total:
        print(f"等待 {DELAY_BETWEEN_FILES_SECONDS // 60} 分钟后继续处理下一个 PDF...")
        time.sleep(DELAY_BETWEEN_FILES_SECONDS)


# ============================================================
# PDF24：PDF 转 Word
# ============================================================

def pdf24_pdf_to_docx(input_pdf: Path, output_docx: Path):
    print(f"\nPDF24：PDF 转 Word：{input_pdf.name}")

    open_pdf24()
    go_home()

    pyautogui.click(HOME_PDF_TO_OTHER_X, HOME_PDF_TO_OTHER_Y)
    wait(2)

    pyautogui.click(SELECT_FILE_X, SELECT_FILE_Y)
    windows_open_file_dialog(input_pdf)

    wait(3)

    print("正在点击左下角 Text(.txt) 格式下拉框...")
    pyautogui.click(FORMAT_DROPDOWN_X, FORMAT_DROPDOWN_Y)
    wait(1)

    print("正在选择 Word(.docx)...")
    pyautogui.click(WORD_DOCX_OPTION_X, WORD_DOCX_OPTION_Y)
    wait(1)

    print("正在点击 PDF 转 Word 橙色转换按钮...")
    pyautogui.click(PDF_TO_WORD_ORANGE_CONVERT_X, PDF_TO_WORD_ORANGE_CONVERT_Y)

    wait(12)

    clicked = force_click_pdf24_save_button(
        PDF_TO_WORD_SAVE_BUTTON_X,
        PDF_TO_WORD_SAVE_BUTTON_Y,
        description="PDF 转 Word 蓝色保存按钮",
    )

    if not clicked:
        close_pdf24()
        raise RuntimeError("PDF 转 Word：蓝色保存按钮点击后没有弹出另存为窗口")

    windows_save_as_dialog(output_docx)

    if not wait_for_file_exists(output_docx, timeout=60):
        close_pdf24()
        raise FileNotFoundError(f"PDF24 保存英文 Word 超时：{output_docx}")

    close_pdf24()

    print(f"PDF24：英文 Word 已保存到：{output_docx}")


# ============================================================
# PDF24：Word 转 PDF
# ============================================================

def pdf24_docx_to_pdf(input_docx: Path, output_pdf: Path):
    print(f"\nPDF24：中文 Word 转 PDF：{input_docx.name}")

    open_pdf24()
    go_home()

    pyautogui.click(HOME_CONVERT_TO_PDF_X, HOME_CONVERT_TO_PDF_Y)
    wait(2)

    pyautogui.click(SELECT_FILE_X, SELECT_FILE_Y)
    windows_open_file_dialog(input_docx)

    wait(4)

    print("正在点击 Word 转 PDF 橙色转换按钮...")
    pyautogui.click(WORD_TO_PDF_ORANGE_CONVERT_X, WORD_TO_PDF_ORANGE_CONVERT_Y)

    wait(12)

    clicked = force_click_pdf24_save_button(
        WORD_TO_PDF_SAVE_BUTTON_X,
        WORD_TO_PDF_SAVE_BUTTON_Y,
        description="Word 转 PDF 蓝色保存按钮",
    )

    if not clicked:
        close_pdf24()
        raise RuntimeError("Word 转 PDF：蓝色保存按钮点击后没有弹出另存为窗口")

    windows_save_as_dialog(output_pdf)

    if not wait_for_file_exists(output_pdf, timeout=60):
        close_pdf24()
        raise FileNotFoundError(f"PDF24 保存中文 PDF 超时：{output_pdf}")

    close_pdf24()

    print(f"PDF24：中文 PDF 已保存到：{output_pdf}")


# ============================================================
# Argos 离线翻译
# ============================================================

def is_translatable_text(text: str) -> bool:
    if not text:
        return False

    stripped = text.strip()

    if not stripped:
        return False

    if not re.search(r"[A-Za-z]", stripped):
        return False

    if re.match(r"^https?://", stripped, re.I):
        return False

    if re.match(r"^[\w\.-]+@[\w\.-]+\.\w+$", stripped):
        return False

    if re.match(r"^[\d\.\-\(\)\[\]\s]+$", stripped):
        return False

    return True


class TranslatorWithCache:
    def __init__(self, source_lang="en", target_lang="zh", cache_path=None):
        self.source_lang = source_lang
        self.target_lang = target_lang
        self.cache_path = Path(cache_path) if cache_path else None
        self.cache = {}

        if self.cache_path and self.cache_path.exists():
            try:
                self.cache = json.loads(self.cache_path.read_text(encoding="utf-8"))
            except Exception:
                self.cache = {}

        self.check_argos_model()

    def check_argos_model(self):
        installed_languages = argostranslate.translate.get_installed_languages()

        from_lang = None
        to_lang = None

        for lang in installed_languages:
            if lang.code == self.source_lang:
                from_lang = lang
            if lang.code == self.target_lang:
                to_lang = lang

        if from_lang is None or to_lang is None:
            raise RuntimeError(
                "没有找到 Argos English -> Chinese 离线模型。\n"
                "请确认你已经安装 en -> zh 模型。\n"
                "可以先运行 check_argos_model.py 测试。"
            )

        translation = from_lang.get_translation(to_lang)

        if translation is None:
            raise RuntimeError("Argos 已安装语言，但没有找到 en -> zh 翻译方向。")

        print("Argos 离线翻译模型检查通过：en -> zh")

    def save_cache(self):
        if self.cache_path:
            self.cache_path.parent.mkdir(parents=True, exist_ok=True)
            self.cache_path.write_text(
                json.dumps(self.cache, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )

    def translate(self, text: str) -> str:
        if not is_translatable_text(text):
            return text

        key = text.strip()

        if key in self.cache:
            return self.cache[key]

        for attempt in range(3):
            try:
                translated = argostranslate.translate.translate(
                    key,
                    self.source_lang,
                    self.target_lang,
                )

                if not translated:
                    translated = key

                translated = str(translated).strip()

                self.cache[key] = translated
                self.save_cache()

                return translated

            except Exception as e:
                print(f"    Argos 翻译失败，重试 {attempt + 1}/3：{e}")
                time.sleep(1)

        print("    Argos 多次失败，保留原文：", key[:80])
        return key


# ============================================================
# DOCX XML 处理
# ============================================================

def should_process_xml(name: str) -> bool:
    if not name.endswith(".xml"):
        return False

    return any(name.startswith(prefix) for prefix in WORD_XML_PREFIXES)


def get_paragraph_text(paragraph):
    text_nodes = paragraph.xpath(".//w:t", namespaces=NS)
    return "".join(t.text or "" for t in text_nodes), text_nodes


def replace_paragraph_text(text_nodes, new_text):
    if not text_nodes:
        return

    text_nodes[0].text = new_text
    text_nodes[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

    for t in text_nodes[1:]:
        t.text = ""


def translate_xml_content(xml_bytes, translator: TranslatorWithCache, file_name: str):
    parser = etree.XMLParser(remove_blank_text=False, recover=True)
    root = etree.fromstring(xml_bytes, parser)

    paragraphs = root.xpath(".//w:p", namespaces=NS)
    changed = False

    total = len(paragraphs)
    translated_count = 0

    for idx, p in enumerate(paragraphs, start=1):
        original_text, text_nodes = get_paragraph_text(p)

        if not is_translatable_text(original_text):
            continue

        translated_text = translator.translate(original_text)
        translated_count += 1

        if translated_text and translated_text != original_text:
            replace_paragraph_text(text_nodes, translated_text)
            changed = True

        if translated_count % 20 == 0:
            print(f"    {file_name}：已翻译 {translated_count} 个段落，当前段落 {idx}/{total}")

    if translated_count > 0:
        print(f"    {file_name}：共翻译 {translated_count} 个段落")

    if not changed:
        return xml_bytes

    return etree.tostring(
        root,
        xml_declaration=True,
        encoding="UTF-8",
        standalone=False,
    )


def translate_docx(input_docx: Path, output_docx: Path, translator: TranslatorWithCache):
    output_docx.parent.mkdir(parents=True, exist_ok=True)

    temp_docx = output_docx.with_suffix(".tmp.docx")

    if temp_docx.exists():
        temp_docx.unlink()

    print(f"\nPython + Argos：翻译 Word：{input_docx.name}")

    with zipfile.ZipFile(input_docx, "r") as zin:
        with zipfile.ZipFile(temp_docx, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)

                if should_process_xml(item.filename):
                    try:
                        data = translate_xml_content(data, translator, item.filename)
                    except Exception as e:
                        print(f"  警告：{item.filename} 处理失败，保留原内容：{e}")

                zout.writestr(item, data)

    shutil.move(str(temp_docx), str(output_docx))

    print(f"Python + Argos：中文 Word 完成：{output_docx}")


# ============================================================
# 文件夹结构处理
# ============================================================

def collect_pdfs(input_root: Path, limit: int):
    pdfs = sorted(input_root.rglob("*.pdf"))

    if limit and limit > 0:
        return pdfs[:limit]

    return pdfs


def process_all(input_root: Path, output_root: Path, overwrite: bool, limit: int):
    input_root = input_root.resolve()
    output_root = output_root.resolve()
    wd_root = Path(WD_ROOT).resolve()

    if not input_root.exists():
        raise FileNotFoundError(f"输入文件夹不存在：{input_root}")

    output_root.mkdir(parents=True, exist_ok=True)
    wd_root.mkdir(parents=True, exist_ok=True)

    pdf_files = collect_pdfs(input_root, limit)

    if not pdf_files:
        print("没有找到 PDF 文件")
        return

    print(f"本次准备处理 {len(pdf_files)} 个 PDF")
    print(f"输入目录：{input_root}")
    print(f"Word 中转目录：{wd_root}")
    print(f"输出 PDF 目录：{output_root}")

    cache_path = output_root / "translation_cache_argos.json"

    translator = TranslatorWithCache(
        source_lang="en",
        target_lang="zh",
        cache_path=cache_path,
    )

    success = 0
    failed = 0

    for index, input_pdf in enumerate(pdf_files, start=1):
        print("\n" + "=" * 80)
        print(f"处理 {index}/{len(pdf_files)}：{input_pdf}")

        relative_pdf = input_pdf.relative_to(input_root)
        output_pdf = output_root / relative_pdf

        wd_subdir = wd_root / relative_pdf.parent
        wd_subdir.mkdir(parents=True, exist_ok=True)

        en_docx = wd_subdir / f"{input_pdf.stem}_en.docx"
        zh_docx = wd_subdir / f"{input_pdf.stem}_zh.docx"

        try:
            if output_pdf.exists() and not overwrite:
                print(f"最终中文 PDF 已存在，跳过：{output_pdf}")
                success += 1
                wait_between_files(index, len(pdf_files))
                continue

            if not en_docx.exists() or overwrite:
                pdf24_pdf_to_docx(input_pdf, en_docx)
            else:
                print(f"英文 Word 已存在，跳过 PDF 转 Word：{en_docx}")

            if not en_docx.exists():
                raise FileNotFoundError(f"PDF24 没有成功生成英文 Word：{en_docx}")

            if not zh_docx.exists() or overwrite:
                translate_docx(en_docx, zh_docx, translator)
            else:
                print(f"中文 Word 已存在，跳过翻译：{zh_docx}")

            if not zh_docx.exists():
                raise FileNotFoundError(f"中文 Word 没有生成：{zh_docx}")

            pdf24_docx_to_pdf(zh_docx, output_pdf)

            if not output_pdf.exists():
                raise FileNotFoundError(f"PDF24 没有成功生成中文 PDF：{output_pdf}")

            print(f"完成：{output_pdf}")
            success += 1

            wait_between_files(index, len(pdf_files))

        except Exception as e:
            failed += 1

            print("\n" + "!" * 80)
            print("当前 PDF 处理失败")
            print(f"失败文件位置：{input_pdf}")
            print(f"预计输出位置：{output_pdf}")
            print(f"错误原因：{e}")
            print(f"当前失败数量：{failed}/{MAX_FAILED_FILES}")
            print("!" * 80)

            close_pdf24()

            if failed > MAX_FAILED_FILES:
                print(f"失败文件数量已经超过 {MAX_FAILED_FILES} 个，脚本停止继续处理。")
                print("\n" + "=" * 80)
                print("提前结束")
                print(f"成功：{success}")
                print(f"失败：{failed}")
                print(f"Word 中转目录：{wd_root}")
                print(f"输出 PDF 目录：{output_root}")
                return

            wait_between_files(index, len(pdf_files))

    print("\n" + "=" * 80)
    print("全部结束")
    print(f"成功：{success}")
    print(f"失败：{failed}")
    print(f"Word 中转目录：{wd_root}")
    print(f"输出 PDF 目录：{output_root}")


def main():
    parser = argparse.ArgumentParser(
        description="使用 PDF24 完成 PDF/Word 格式转换，并用 Argos 离线翻译 PDF"
    )

    parser.add_argument(
        "--input",
        default=DEFAULT_INPUT,
        help="输入 PDF 文件夹",
    )

    parser.add_argument(
        "--output",
        default=DEFAULT_OUTPUT,
        help="输出中文 PDF 文件夹",
    )

    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="如果输出文件已存在，强制重新生成",
    )

    parser.add_argument(
        "--limit",
        type=int,
        default=0,
        help="限制处理数量。默认 0 表示处理全部"
    )

    args = parser.parse_args()

    process_all(
        input_root=Path(args.input),
        output_root=Path(args.output),
        overwrite=args.overwrite,
        limit=args.limit,
    )


if __name__ == "__main__":
    main()