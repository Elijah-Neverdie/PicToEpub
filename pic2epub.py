import os
import sys
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
from pathlib import Path
import shutil
import tempfile
import re
from zipfile import ZipFile, ZIP_STORED
from lxml import etree
import time


# ========== 精简的多语言字典 ==========
LANGUAGES = {
    "中文": {
        "window_title": "Pic2EPUB",
        "browse": "浏览...",
        "confirm": "生成EPUB",
        "cancel": "取消",
        "folder_label": "图片文件夹路径:",
        "processing": "正在生成EPUB...",
        "scanning": "扫描图片文件中...",
        "packaging": "正在打包EPUB文件...",
        "progress_overall": "总进度: {done}/{total}",
        "progress_current": "当前: {current}/{total} ({percent}%)",
        "cancelling": "正在取消...",
        "error_invalid_folder": "请选择一个有效的文件夹！",
        "warning_no_images": "文件夹中没有可识别的图片！",
        "error_conversion_failed": "生成失败：\n{error}",
        "success_epub_created": "EPUB 已生成：\n{output}",
        "batch_prompt": "检测到 {count} 个子文件夹。请选择处理方式：",
        "batch_option_separate": "为每个子文件夹单独生成 EPUB",
        "batch_option_merge": "将所有图片合并到一个 EPUB",
        "batch_option_cancel": "取消操作",
        "overwrite_title": "文件已存在",
        "overwrite_message": "文件 '{file}' 已存在。您想要怎么做？",
        "overwrite_skip": "跳过",
        "overwrite_overwrite": "覆盖",
        "overwrite_cancel": "取消",
        "overwrite_apply_all": "应用于所有"
    },
    "English": {
        "window_title": "Pic2EPUB",
        "browse": "Browse...",
        "confirm": "Generate EPUB",
        "cancel": "Cancel",
        "folder_label": "Image Folder Path:",
        "processing": "Generating EPUB...",
        "scanning": "Scanning image files...",
        "packaging": "Packaging EPUB file...",
        "progress_overall": "Overall: {done}/{total}",
        "progress_current": "Current: {current}/{total} ({percent}%)",
        "cancelling": "Cancelling...",
        "error_invalid_folder": "Please select a valid folder!",
        "warning_no_images": "No recognizable images in the folder!",
        "error_conversion_failed": "Generation failed:\n{error}",
        "success_epub_created": "EPUB created:\n{output}",
        "batch_prompt": "Detected {count} subfolders. Please select processing method:",
        "batch_option_separate": "Create separate EPUB for each subfolder",
        "batch_option_merge": "Merge all images into one EPUB",
        "batch_option_cancel": "Cancel operation",
        "overwrite_title": "File Exists",
        "overwrite_message": "File '{file}' already exists. What would you like to do?",
        "overwrite_skip": "Skip",
        "overwrite_overwrite": "Overwrite",
        "overwrite_cancel": "Cancel",
        "overwrite_apply_all": "Apply to all"
    }
}


class OverwriteDialog:
    """覆盖确认对话框"""
    def __init__(self, parent, filename, lang="中文"):
        self.result = None
        self.apply_all = False
        
        tr = lambda k: LANGUAGES[lang].get(k, k)
        
        dialog = tk.Toplevel(parent)
        dialog.title(tr("overwrite_title"))
        dialog.resizable(False, False)
        dialog.transient(parent)
        dialog.grab_set()
        
        # 居中显示
        parent.update_idletasks()
        px = parent.winfo_rootx()
        py = parent.winfo_rooty()
        pw = parent.winfo_width()
        ph = parent.winfo_height()
        dx = px + (pw - 400) // 2
        dy = py + (ph - 200) // 2
        dialog.geometry(f"+{dx}+{dy}")
        
        # 消息
        msg_label = tk.Label(
            dialog, 
            text=tr("overwrite_message").format(file=filename),
            font=("微软雅黑", 10),
            wraplength=350,
            justify="left"
        )
        msg_label.pack(pady=20)
        
        # 复选框
        self.apply_var = tk.BooleanVar()
        apply_check = tk.Checkbutton(
            dialog, 
            text=tr("overwrite_apply_all"),
            variable=self.apply_var,
            font=("微软雅黑", 9)
        )
        apply_check.pack(pady=10)
        
        # 按钮框架
        btn_frame = tk.Frame(dialog)
        btn_frame.pack(pady=10)
        
        def set_result(value):
            self.result = value
            self.apply_all = self.apply_var.get()
            dialog.destroy()
        
        tk.Button(
            btn_frame, 
            text=tr("overwrite_skip"),
            command=lambda: set_result("skip"),
            width=10
        ).pack(side=tk.LEFT, padx=5)
        
        tk.Button(
            btn_frame,
            text=tr("overwrite_overwrite"),
            command=lambda: set_result("overwrite"),
            width=10
        ).pack(side=tk.LEFT, padx=5)
        
        tk.Button(
            btn_frame,
            text=tr("overwrite_cancel"),
            command=lambda: set_result("cancel"),
            width=10
        ).pack(side=tk.LEFT, padx=5)
        
        dialog.wait_window()


def extract_number(filename):
    """从文件名中提取数字用于排序"""
    match = re.search(r'\d+', filename)
    return int(match.group()) if match else float('inf')


def get_cover_file(image_files):
    """选择最佳的cover文件"""
    cover_candidates = []
    
    for f in image_files:
        if 'cover' in f.lower():
            # 优先选择jpg文件
            if f.lower().endswith(('.jpg', '.jpeg')):
                cover_candidates.append((f, 2))  # 优先级2
            else:
                cover_candidates.append((f, 1))  # 优先级1
    
    if not cover_candidates:
        return None
    
    # 排序：先按优先级
    cover_candidates.sort(key=lambda x: -x[1])
    return cover_candidates[0][0]


def sort_image_files(image_files):
    """排序图片文件：cover文件在前，其余按数字排序"""
    cover_file = get_cover_file(image_files)
    others = []
    
    for f in image_files:
        if f != cover_file:
            others.append(f)
    
    # 对剩余文件按数字排序
    others_sorted = sorted(others, key=lambda x: (extract_number(x), x))
    
    if cover_file:
        return [cover_file] + others_sorted
    else:
        return others_sorted


def check_and_install_deps():
    """检查并安装依赖"""
    try:
        from PIL import Image
        import lxml.etree
    except ImportError as e:
        missing = str(e).split()[-1].strip("'")
        if missing == "lxml":
            missing = "lxml"
        elif missing == "PIL":
            missing = "Pillow"
            
        msg = LANGUAGES["中文"]["dependency_missing"] if "中文" in LANGUAGES else f"Missing dependency: {missing}"
        if messagebox.askyesno("Dependency Missing", msg):
            try:
                if missing == "lxml":
                    subprocess.check_call([sys.executable, "-m", "pip", "install", "lxml"])
                else:
                    subprocess.check_call([sys.executable, "-m", "pip", "install", "Pillow"])
                messagebox.showinfo("Success", LANGUAGES["中文"]["dependency_installed"] if "中文" in LANGUAGES else "Dependencies installed. Please restart.")
                sys.exit(0)
            except Exception as install_err:
                messagebox.showerror("Install Failed", str(install_err))
                sys.exit(1)
        else:
            sys.exit(0)


def get_supported_image_extensions():
    """获取支持的图片格式"""
    from PIL import Image
    exts = Image.registered_extensions()
    return {ext.lower() for ext in exts.keys()}


def scan_images(folder, progress_callback=None, stop_event=None):
    """扫描文件夹中的图片文件，忽略EPUB文件"""
    supported = get_supported_image_extensions()
    try:
        # 过滤掉EPUB文件
        all_files = [f for f in os.listdir(folder) 
                    if os.path.isfile(os.path.join(folder, f)) 
                    and not f.lower().endswith('.epub')]
    except Exception:
        return [], []
    
    image_files = []
    
    for i, f in enumerate(all_files):
        # 检查是否取消
        if stop_event and stop_event.is_set():
            raise InterruptedError("User cancelled")
            
        if Path(f).suffix.lower() in supported:
            image_files.append(f)
        
        # 每处理一个文件就更新进度
        if progress_callback:
            progress_callback(i + 1, len(all_files))
    
    return image_files, []


def get_image_media_type(filename):
    """根据文件扩展名获取媒体类型"""
    ext = Path(filename).suffix.lower()
    if ext in ['.jpg', '.jpeg']:
        return 'image/jpeg'
    elif ext == '.png':
        return 'image/png'
    elif ext == '.gif':
        return 'image/gif'
    elif ext == '.webp':
        return 'image/webp'
    elif ext == '.svg':
        return 'image/svg+xml'
    else:
        return 'image/jpeg'  # 默认


def create_epub_from_images(image_paths, output_file, book_title, progress_callback=None, stop_event=None):
    """从图片列表创建EPUB文件"""
    # 确保临时目录存在
    os.makedirs('META-INF', exist_ok=True)
    os.makedirs('OEBPS/images', exist_ok=True)
    os.makedirs('OEBPS/text', exist_ok=True)

    # 创建mimetype文件
    with open('mimetype', 'w') as f:
        f.write('application/epub+zip')

    # 创建container.xml
    container = etree.Element('container', version='1.0', xmlns='urn:oasis:names:tc:opendocument:xmlns:container')
    rootfiles = etree.SubElement(container, 'rootfiles')
    etree.SubElement(rootfiles, 'rootfile', **{
        'full-path': 'OEBPS/content.opf',
        'media-type': 'application/oebps-package+xml'
    })
    with open('META-INF/container.xml', 'wb') as f:
        f.write(etree.tostring(container, pretty_print=True, xml_declaration=True, encoding='UTF-8'))
    
    if progress_callback:
        progress_callback(1, len(image_paths) + 10)  # 基础步骤

    # 创建content.opf
    opf = etree.Element('package', version='3.0', xmlns='http://www.idpf.org/2007/opf', unique_identifier='bookid')
    metadata = etree.SubElement(opf, 'metadata', nsmap={
        'dc': 'http://purl.org/dc/elements/1.1/',
        'opf': 'http://www.idpf.org/2007/opf'
    })
    etree.SubElement(metadata, '{http://purl.org/dc/elements/1.1/}identifier', id='bookid').text = 'urn:uuid:1234567890'
    etree.SubElement(metadata, '{http://purl.org/dc/elements/1.1/}title').text = book_title
    etree.SubElement(metadata, '{http://purl.org/dc/elements/1.1/}language').text = 'zh'

    manifest = etree.SubElement(opf, 'manifest')
    spine = etree.SubElement(opf, 'spine', toc='ncx')
    etree.SubElement(manifest, 'item', id='ncx', href='toc.ncx', media_type='application/x-dtbncx+xml')

    # 创建NCX目录
    ncx = etree.Element('ncx', xmlns='http://www.daisy.org/z3986/2005/ncx/', version='2005-1')
    head = etree.SubElement(ncx, 'head')
    for name, content in [('dtb:uid', 'urn:uuid:1234567890'), ('dtb:depth', '1'), ('dtb:totalPageCount', '0'), ('dtb:maxPageNumber', '0')]:
        etree.SubElement(head, 'meta', name=name, content=content)
    doc_title = etree.SubElement(ncx, 'docTitle')
    etree.SubElement(doc_title, 'text').text = book_title
    nav_map = etree.SubElement(ncx, 'navMap')
    
    if progress_callback:
        progress_callback(2, len(image_paths) + 10)  # 元数据创建完成

    # 添加所有图片页面 - 每个图片处理都会更新进度
    total_images = len(image_paths)
    for i, img_path in enumerate(image_paths):
        # 检查是否取消
        if stop_event and stop_event.is_set():
            raise InterruptedError("User cancelled")
            
        img_filename = f"img_{i:04d}{Path(img_path).suffix}"
        dest_img = f'OEBPS/images/{img_filename}'
        shutil.copy2(img_path, dest_img)
        
        media_type = get_image_media_type(img_path)
        etree.SubElement(manifest, 'item', id=f'img{i}', href=f'images/{img_filename}', media_type=media_type)
        
        # 创建XHTML页面
        xhtml = etree.Element('html', xmlns='http://www.w3.org/1999/xhtml')
        head = etree.SubElement(xhtml, 'head')
        etree.SubElement(head, 'title').text = f'Page {i+1}'
        body = etree.SubElement(xhtml, 'body')
        etree.SubElement(body, 'img', src=f'../images/{img_filename}', style='width: 100%; height: auto;')
        
        xhtml_path = f'OEBPS/text/page_{i:04d}.xhtml'
        with open(xhtml_path, 'wb') as f:
            f.write(etree.tostring(xhtml, pretty_print=True, xml_declaration=True, encoding='UTF-8'))
        
        etree.SubElement(manifest, 'item', id=f'page{i}', href=f'text/page_{i:04d}.xhtml', media_type='application/xhtml+xml')
        etree.SubElement(spine, 'itemref', idref=f'page{i}')
        
        # 添加导航点
        nav_point = etree.SubElement(nav_map, 'navPoint', id=f'navPoint-{i+1}', playOrder=str(i+1))
        label = etree.SubElement(nav_point, 'navLabel')
        etree.SubElement(label, 'text').text = f'Page {i+1}'
        etree.SubElement(nav_point, 'content', src=f'text/page_{i:04d}.xhtml')
        
        # 每处理一张图片就更新进度
        if progress_callback:
            progress_callback(3 + i, len(image_paths) + 10)

    if progress_callback:
        progress_callback(3 + total_images, len(image_paths) + 10)  # 所有页面创建完成

    # 保存OPF和NCX文件
    with open('OEBPS/content.opf', 'wb') as f:
        f.write(etree.tostring(opf, pretty_print=True, xml_declaration=True, encoding='UTF-8'))
    with open('OEBPS/toc.ncx', 'wb') as f:
        f.write(etree.tostring(ncx, pretty_print=True, xml_declaration=True, encoding='UTF-8'))
    
    if progress_callback:
        progress_callback(4 + total_images, len(image_paths) + 10)  # 内容文件保存完成

    # 打包EPUB - 这是最耗时的部分，我们将其分解为多个步骤
    if progress_callback:
        progress_callback(5 + total_images, len(image_paths) + 10)  # 开始打包
    
    # 收集所有需要打包的文件
    all_files = []
    all_files.append(('mimetype', 'mimetype'))
    
    # META-INF 文件
    for root, _, files in os.walk('META-INF'):
        for f in files:
            full_path = os.path.join(root, f)
            rel_path = os.path.relpath(full_path, '.')
            all_files.append((full_path, rel_path))
    
    # OEBPS 文件
    for root, _, files in os.walk('OEBPS'):
        for f in files:
            full_path = os.path.join(root, f)
            rel_path = os.path.relpath(full_path, '.')
            all_files.append((full_path, rel_path))
    
    # 分批打包文件，每批更新一次进度
    with ZipFile(output_file, 'w') as epub:
        # 先添加mimetype（必须无压缩）
        epub.write('mimetype', compress_type=ZIP_STORED)
        
        # 添加其他文件，分批处理
        batch_size = max(1, len(all_files) // 10)  # 分成10批左右
        for i in range(0, len(all_files), batch_size):
            # 检查是否取消
            if stop_event and stop_event.is_set():
                raise InterruptedError("User cancelled")
                
            batch = all_files[i:i + batch_size]
            for full_path, rel_path in batch:
                if rel_path != 'mimetype':  # mimetype已经单独处理
                    epub.write(full_path, rel_path)
            
            # 更新打包进度
            if progress_callback:
                progress_value = 5 + total_images + int((i + len(batch)) / len(all_files) * 4)
                progress_callback(progress_value, len(image_paths) + 10)
    
    if progress_callback:
        progress_callback(9 + total_images, len(image_paths) + 10)  # EPUB打包完成

    # 清理临时文件
    if os.path.exists('mimetype'):
        os.remove('mimetype')
    if os.path.exists('META-INF'):
        shutil.rmtree('META-INF')
    if os.path.exists('OEBPS'):
        shutil.rmtree('OEBPS')
    
    if progress_callback:
        progress_callback(10 + total_images, len(image_paths) + 10)  # 100% - 完成


def get_valid_subfolders(base_folder, progress_callback=None, stop_event=None):
    """返回包含至少一张图片的子文件夹列表"""
    subfolders = []
    try:
        items = [item for item in os.listdir(base_folder) 
                if os.path.isdir(os.path.join(base_folder, item))]
        
        for i, item in enumerate(items):
            # 检查是否取消
            if stop_event and stop_event.is_set():
                raise InterruptedError("User cancelled")
                
            path = os.path.join(base_folder, item)
            if os.path.isdir(path):
                imgs, _ = scan_images(path, stop_event=stop_event)
                if imgs:
                    subfolders.append(path)
            
            if progress_callback:
                progress_callback(i + 1, len(items))
    except Exception:
        pass
    return subfolders


def get_all_images_from_subfolders(base_folder, progress_callback=None, stop_event=None):
    """从所有子文件夹获取图片文件"""
    all_images = []
    subfolders = get_valid_subfolders(base_folder, stop_event=stop_event)
    
    # 收集所有图片
    total_folders = len(subfolders)
    for folder_idx, folder in enumerate(subfolders):
        # 检查是否取消
        if stop_event and stop_event.is_set():
            raise InterruptedError("User cancelled")
            
        imgs, _ = scan_images(folder, stop_event=stop_event)
        for img_idx, img in enumerate(imgs):
            all_images.append(os.path.join(folder, img))
            
            # 更新进度
            if progress_callback:
                current = folder_idx * 100 + img_idx + 1
                total = total_folders * 100  # 估计值，确保进度平滑
                progress_callback(current, total)
    
    return all_images


class ScanProgressWindow:
    """扫描进度窗口"""
    def __init__(self, parent, lang="中文"):
        self.window = tk.Toplevel(parent)
        self.lang = lang
        tr = lambda k: LANGUAGES[lang].get(k, k)
        self.window.title(tr("scanning"))
        self.window.resizable(False, False)
        self.window.transient(parent)
        self.window.grab_set()

        parent.update_idletasks()
        px = parent.winfo_rootx()
        py = parent.winfo_rooty()
        pw = parent.winfo_width()
        ph = parent.winfo_height()
        dx = px + (pw - 400) // 2
        dy = py + (ph - 120) // 2
        self.window.geometry(f"+{dx}+{dy}")

        self.parent = parent
        self.stop_event = threading.Event()
        self._closed = False

        # 扫描进度
        self.scan_label = tk.Label(self.window, text=tr("scanning"), font=("微软雅黑", 10))
        self.scan_label.pack(pady=(20, 10))
        
        self.scan_bar = ttk.Progressbar(self.window, length=350, mode='determinate')
        self.scan_bar.pack(pady=(0, 20))

        self.cancel_btn = tk.Button(self.window, text=tr("cancel"), command=self.cancel, width=10)
        self.cancel_btn.pack(pady=(0, 15))

        self.window.protocol("WM_DELETE_WINDOW", self.on_close)

    def on_close(self):
        self.cancel()
        self.window.destroy()

    def cancel(self):
        self.stop_event.set()
        tr = lambda k: LANGUAGES[self.lang].get(k, k)
        self.scan_label.config(text=tr("cancelling"))
        self.cancel_btn.config(state="disabled")

    def update_scan(self, current, total):
        if self._closed:
            return
        percent = int(current / total * 100) if total > 0 else 0
        self.scan_bar['maximum'] = total
        self.scan_bar['value'] = current
        
        tr = lambda k: LANGUAGES[self.lang].get(k, k)
        text = f"{tr('scanning')} ({current}/{total})"
        self.scan_label.config(text=text)
        self.window.update_idletasks()

    def close(self):
        self._closed = True
        self.window.destroy()


class ProgressWindow:
    def __init__(self, parent, is_batch=False, total_books=1, lang="中文"):
        self.window = tk.Toplevel(parent)
        self.lang = lang
        tr = lambda k: LANGUAGES[lang].get(k, k)
        self.window.title(tr("processing"))
        self.window.resizable(False, False)
        self.window.transient(parent)
        self.window.grab_set()

        parent.update_idletasks()
        px = parent.winfo_rootx()
        py = parent.winfo_rooty()
        pw = parent.winfo_width()
        ph = parent.winfo_height()
        dx = px + (pw - 450) // 2  # 增加宽度以容纳动画
        dy = py + (ph - (220 if is_batch else 160)) // 2  # 增加高度以容纳书名
        self.window.geometry(f"+{dx}+{dy}")

        self.parent = parent
        self.stop_event = threading.Event()
        self._closed = False
        self.animation_running = False
        self.animation_index = 0
        self.animation_chars = ["⠋", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧", "⠇", "⠏"]
        self.current_book = ""

        # 总进度（仅批量）
        if is_batch:
            self.overall_label = tk.Label(self.window, text=tr("progress_overall").format(done=0, total=total_books), font=("微软雅黑", 9))
            self.overall_label.pack(pady=(10, 0))
            self.overall_bar = ttk.Progressbar(self.window, length=400, mode='determinate', maximum=total_books)
            self.overall_bar.pack(pady=(2, 5))

        # 当前书籍名称显示
        self.book_label = tk.Label(self.window, text="", font=("微软雅黑", 10, "bold"), wraplength=400)
        self.book_label.pack(pady=(5, 0))
        
        # 当前书籍进度
        self.current_label = tk.Label(self.window, text=tr("processing"), font=("微软雅黑", 9))
        self.current_label.pack()
        self.current_bar = ttk.Progressbar(self.window, length=400, mode='determinate')
        self.current_bar.pack(pady=(2, 5))
        
        # 添加动画指示器
        self.animation_label = tk.Label(self.window, text="", font=("微软雅黑", 12))
        self.animation_label.pack(pady=(0, 10))

        self.cancel_btn = tk.Button(self.window, text=tr("cancel"), command=self.cancel, width=10)
        self.cancel_btn.pack(pady=(0, 10))

        self.window.protocol("WM_DELETE_WINDOW", self.on_close)
        
        # 开始动画
        self.start_animation()

    def start_animation(self):
        """开始动画显示"""
        self.animation_running = True
        self.update_animation()
    
    def stop_animation(self):
        """停止动画显示"""
        self.animation_running = False
    
    def update_animation(self):
        """更新动画帧"""
        if not self.animation_running or self._closed:
            return
            
        tr = lambda k: LANGUAGES[self.lang].get(k, k)
        
        # 检查是否在打包阶段
        if "打包" in self.current_label.cget("text") or "Packaging" in self.current_label.cget("text"):
            # 在打包阶段显示更明显的动画
            self.animation_index = (self.animation_index + 1) % len(self.animation_chars)
            animation_char = self.animation_chars[self.animation_index]
            self.animation_label.config(text=f"{animation_char} {tr('packaging')} {animation_char}")
        else:
            # 在其他阶段显示简单的动画
            self.animation_index = (self.animation_index + 1) % len(self.animation_chars)
            self.animation_label.config(text=self.animation_chars[self.animation_index])
        
        # 每100毫秒更新一次动画
        self.window.after(100, self.update_animation)
        
    def set_current_book(self, book_name):
        """设置当前正在处理的书籍名称"""
        self.current_book = book_name
        # 如果书名超过20个字符，则截断并添加省略号
        if len(book_name) > 20:
            display_name = book_name[:17] + '...'
        else:
            display_name = book_name
        self.book_label.config(text=f"《{display_name}》")

    def on_close(self):
        self.cancel()
        self.window.destroy()

    def cancel(self):
        self.stop_event.set()
        tr = lambda k: LANGUAGES[self.lang].get(k, k)
        self.current_label.config(text=tr("cancelling"))
        self.cancel_btn.config(state="disabled")
        self.stop_animation()

    def update_overall(self, done, total):
        if hasattr(self, 'overall_label'):
            tr = lambda k: LANGUAGES[self.lang].get(k, k)
            self.overall_label.config(text=tr("progress_overall").format(done=done, total=total))
            self.overall_bar['value'] = done
            self.window.update_idletasks()

    def update_current(self, current, total):
        if self._closed:
            return
            
        percent = int(current / total * 100) if total > 0 else 0
        self.current_bar['maximum'] = total
        self.current_bar['value'] = current
        
        tr = lambda k: LANGUAGES[self.lang].get(k, k)
        
        # 根据进度显示不同的状态信息
        if percent >= 95:
            status_text = tr("packaging")
        else:
            status_text = tr("processing")
            
        text = f"{status_text} - {tr('progress_current').format(current=current, total=total, percent=percent)}"
        self.current_label.config(text=text)
        
        # 强制更新UI
        self.window.update_idletasks()

    def close(self):
        self._closed = True
        self.stop_animation()
        self.window.destroy()


class OverwritePolicy:
    """覆盖策略管理器"""
    def __init__(self):
        self.global_decision = None  # 'skip', 'overwrite', 'cancel'
    
    def should_overwrite(self, parent, filename, lang="中文"):
        """检查是否应该覆盖文件"""
        if self.global_decision:
            return self.global_decision != "skip"
        
        dialog = OverwriteDialog(parent, filename, lang)
        
        if dialog.apply_all:
            self.global_decision = dialog.result
        
        if dialog.result == "cancel":
            raise InterruptedError("User cancelled overwrite dialog")
        
        return dialog.result == "overwrite"


def run_single_conversion(folder, update_current, stop_event, lang="中文", output_dir=None, overwrite_policy=None, progress_win=None):
    """执行单个文件夹的转换"""
    # 设置当前书籍名称
    if progress_win:
        folder_name = os.path.basename(os.path.normpath(folder))
        progress_win.set_current_book(folder_name)
    
    # 扫描图片 - 占30%进度
    def update_scan_progress(current, total):
        # 将扫描进度映射到总进度的0-30%
        mapped_current = int(current / total * 30) if total > 0 else 0
        update_current(mapped_current, 100)
    
    image_files, _ = scan_images(folder, update_scan_progress, stop_event)
    if not image_files:
        raise ValueError(LANGUAGES[lang]["warning_no_images"])

    sorted_image_files = sort_image_files(image_files)
    image_paths = [os.path.join(folder, f) for f in sorted_image_files]
    
    folder_name = os.path.basename(os.path.normpath(folder))
    epub_name = folder_name + ".epub"
    
    # 如果指定了输出目录，则保存到输出目录，否则保存到原文件夹
    if output_dir:
        output_path = os.path.join(output_dir, epub_name)
    else:
        output_path = os.path.join(folder, epub_name)
    
    # 检查文件是否已存在
    if os.path.exists(output_path):
        if overwrite_policy and not overwrite_policy.should_overwrite(overwrite_policy.parent_window, epub_name, lang):
            return None  # 跳过这个文件
    
    # 生成EPUB - 占70%进度
    def update_epub_progress(current, total):
        # 将EPUB生成进度映射到总进度的30-100%
        mapped_current = 30 + int(current / total * 70) if total > 0 else 30
        update_current(mapped_current, 100)
    
    book_title = folder_name
    create_epub_from_images(image_paths, output_path, book_title, update_epub_progress, stop_event)
    return output_path


def run_merged_conversion(base_folder, image_paths, update_current, stop_event, lang="中文", overwrite_policy=None, progress_win=None):
    """执行合并转换（所有子文件夹图片合并到一个EPUB）"""
    if not image_paths:
        raise ValueError(LANGUAGES[lang]["warning_no_images"])

    # 设置当前书籍名称
    if progress_win:
        folder_name = os.path.basename(os.path.normpath(base_folder))
        progress_win.set_current_book(folder_name + " (合并版)")

    folder_name = os.path.basename(os.path.normpath(base_folder))
    epub_name = folder_name + "_merged.epub"
    output_path = os.path.join(base_folder, epub_name)
    
    # 检查文件是否已存在
    if os.path.exists(output_path):
        if overwrite_policy and not overwrite_policy.should_overwrite(overwrite_policy.parent_window, epub_name, lang):
            return None  # 跳过这个文件
    
    book_title = folder_name + " (Merged)"
    create_epub_from_images(image_paths, output_path, book_title, update_current, stop_event)
    return output_path


def run_batch_conversion(folders, progress_win, finish_callback, lang="中文", output_dir=None, overwrite_policy=None):
    """执行批量转换"""
    generated_epubs = []
    total = len(folders)
    try:
        for idx, folder in enumerate(folders):
            if progress_win.stop_event.is_set():
                break
            progress_win.update_overall(idx, total)

            def update_current(c, t):
                progress_win.update_current(c, t)

            try:
                output_path = run_single_conversion(folder, update_current, progress_win.stop_event, lang, output_dir, overwrite_policy, progress_win)
                if output_path:  # 只有当不跳过时才添加到列表
                    generated_epubs.append(output_path)
            except InterruptedError:
                # 用户取消，跳出循环
                break
            except Exception as e:
                finish_callback(success=False, error=str(e), lang=lang)
                return

        progress_win.update_overall(len(generated_epubs), total)
        finish_callback(success=True, generated=generated_epubs, lang=lang)

    except InterruptedError:
        finish_callback(success=False, cancelled=True, generated=generated_epubs, lang=lang)
    except Exception as e:
        finish_callback(success=False, error=str(e), generated=generated_epubs, lang=lang)


def run_merged_batch_conversion(base_folder, progress_win, finish_callback, lang="中文", overwrite_policy=None):
    """执行合并批量转换"""
    try:
        # 获取所有图片 - 使用当前进度条
        def update_scan_progress(current, total):
            progress_win.update_current(current, total)
        
        image_paths = get_all_images_from_subfolders(base_folder, update_scan_progress, progress_win.stop_event)
        if not image_paths:
            raise ValueError(LANGUAGES[lang]["warning_no_images"])
        
        def update_current(c, t):
            progress_win.update_current(c, t)

        # 执行合并转换
        output_path = run_merged_conversion(base_folder, image_paths, update_current, progress_win.stop_event, lang, overwrite_policy, progress_win)
        
        if output_path:  # 只有当不跳过时才添加到列表
            finish_callback(success=True, generated=[output_path], lang=lang)
        else:
            finish_callback(success=True, generated=[], lang=lang)

    except InterruptedError:
        finish_callback(success=False, cancelled=True, generated=[], lang=lang)
    except Exception as e:
        finish_callback(success=False, error=str(e), generated=[], lang=lang)


class App:
    def __init__(self, root):
        self.root = root
        self.current_lang = "中文"
        self.create_widgets()
        self.update_ui_language()

    def tr(self, key):
        return LANGUAGES[self.current_lang].get(key, key)

    def create_widgets(self):
        # 主窗口设置
        self.root.title(self.tr("window_title"))
        self.root.geometry("500x200")
        self.root.resizable(False, False)
        self.root.configure(bg="#f0f0f0")

        # 标题栏
        title_frame = tk.Frame(self.root, bg="#f0f0f0")
        title_frame.pack(fill="x", padx=10, pady=(10, 15))

        self.title_label = tk.Label(title_frame, text=self.tr("window_title"), font=("微软雅黑", 12, "bold"), bg="#f0f0f0")
        self.title_label.pack(side="left")

        self.lang_var = tk.StringVar(value=self.current_lang)
        self.lang_combo = ttk.Combobox(
            title_frame,
            textvariable=self.lang_var,
            values=list(LANGUAGES.keys()),
            state="readonly",
            width=10,
            font=("微软雅黑", 9)
        )
        self.lang_combo.pack(side="right", padx=(10, 0))
        self.lang_combo.bind("<<ComboboxSelected>>", self.on_language_change)

        # 文件夹选择
        folder_frame = tk.Frame(self.root, bg="#f0f0f0")
        folder_frame.pack(fill="x", padx=10, pady=5)

        self.folder_label = tk.Label(folder_frame, text=self.tr("folder_label"), font=("微软雅黑", 9), bg="#f0f0f0")
        self.folder_label.pack(anchor="w")

        entry_frame = tk.Frame(folder_frame, bg="#f0f0f0")
        entry_frame.pack(fill="x", pady=5)

        self.folder_var = tk.StringVar()
        self.entry_path = tk.Entry(entry_frame, textvariable=self.folder_var, width=50, font=("微软雅黑", 9))
        self.entry_path.pack(side="left", fill="x", expand=True)
        
        self.btn_browse = tk.Button(
            entry_frame,
            text=self.tr("browse"),
            command=self.select_folder,
            width=10,
            font=("微软雅黑", 9)
        )
        self.btn_browse.pack(side="right", padx=(5, 0))

        # 生成按钮
        self.btn_convert = tk.Button(
            self.root, 
            text=self.tr("confirm"), 
            command=self.on_convert, 
            width=15, 
            height=2, 
            font=("微软雅黑", 10, "bold"),
            bg="#4CAF50",
            fg="white"
        )
        self.btn_convert.pack(pady=20)

    def on_language_change(self, event=None):
        self.current_lang = self.lang_var.get()
        self.update_ui_language()

    def update_ui_language(self):
        tr = self.tr
        self.root.title(tr("window_title"))
        self.title_label.config(text=tr("window_title"))
        self.folder_label.config(text=tr("folder_label"))
        self.btn_browse.config(text=tr("browse"))
        self.btn_convert.config(text=tr("confirm"))

    def select_folder(self):
        folder = filedialog.askdirectory(title=self.tr("folder_label"))
        if folder:
            self.folder_var.set(folder)

    def scan_subfolders(self):
        """扫描子文件夹（在后台线程中）"""
        folder = self.folder_var.get().strip()
        if not folder or not os.path.isdir(folder):
            return []
        
        self.scan_progress = ScanProgressWindow(self.root, lang=self.current_lang)
        self.scan_result = []
        
        def scan_thread():
            def update_progress(current, total):
                self.root.after(0, lambda: self.scan_progress.update_scan(current, total))
            
            try:
                subfolders = get_valid_subfolders(folder, update_progress, self.scan_progress.stop_event)
                self.root.after(0, lambda: self.on_scan_complete(subfolders))
            except InterruptedError:
                self.root.after(0, lambda: self.on_scan_cancelled())
            except Exception as e:
                self.root.after(0, lambda: self.on_scan_error(str(e)))
        
        thread = threading.Thread(target=scan_thread, daemon=True)
        thread.start()

    def on_scan_complete(self, subfolders):
        """扫描完成回调"""
        self.scan_progress.close()
        self.scan_result = subfolders
        self.process_after_scan()

    def on_scan_cancelled(self):
        """扫描取消回调"""
        self.scan_progress.close()
        # 不显示错误消息，因为这是用户主动取消

    def on_scan_error(self, error):
        """扫描错误回调"""
        self.scan_progress.close()
        messagebox.showerror("Error", f"扫描文件夹时出错：{error}")

    def on_convert(self):
        folder = self.folder_var.get().strip()
        if not folder or not os.path.isdir(folder):
            messagebox.showerror("Error", self.tr("error_invalid_folder"))
            return

        # 先扫描子文件夹
        self.scan_subfolders()

    def process_after_scan(self):
        """扫描完成后处理"""
        folder = self.folder_var.get().strip()
        subfolders = self.scan_result
        
        is_batch = len(subfolders) >= 1

        if is_batch:
            # 创建自定义对话框
            dialog = tk.Toplevel(self.root)
            dialog.title(self.tr("window_title"))
            dialog.resizable(False, False)
            dialog.transient(self.root)
            dialog.grab_set()

            # 居中显示
            self.root.update_idletasks()
            px = self.root.winfo_rootx()
            py = self.root.winfo_rooty()
            pw = self.root.winfo_width()
            ph = self.root.winfo_height()
            dx = px + (pw - 500) // 2
            dy = py + (ph - 200) // 2
            dialog.geometry(f"+{dx}+{dy}")

            # 提示信息
            tr = self.tr
            prompt_label = tk.Label(
                dialog, 
                text=tr("batch_prompt").format(count=len(subfolders)),
                font=("微软雅黑", 10),
                wraplength=450,
                justify="left"
            )
            prompt_label.pack(pady=(20, 15))

            # 按钮框架
            btn_frame = tk.Frame(dialog)
            btn_frame.pack(pady=10)

            def select_separate():
                dialog.destroy()
                self.start_conversion(subfolders, is_batch=True, output_dir=folder)

            def select_merge():
                dialog.destroy()
                self.start_merged_conversion(folder)

            def select_cancel():
                dialog.destroy()

            btn_separate = tk.Button(
                btn_frame, 
                text=tr("batch_option_separate"),
                command=select_separate,
                width=40,
                wraplength=350,
                justify="center",
                font=("微软雅黑", 9)
            )
            btn_separate.pack(pady=5)

            btn_merge = tk.Button(
                btn_frame,
                text=tr("batch_option_merge"),
                command=select_merge,
                width=40,
                wraplength=350,
                justify="center",
                font=("微软雅黑", 9)
            )
            btn_merge.pack(pady=5)

            btn_cancel = tk.Button(
                btn_frame,
                text=tr("batch_option_cancel"),
                command=select_cancel,
                width=20,
                font=("微软雅黑", 9)
            )
            btn_cancel.pack(pady=10)

        else:
            self.start_conversion([folder])

    def start_conversion(self, folders, is_batch=False, output_dir=None):
        """启动单独转换"""
        self.progress_win = ProgressWindow(self.root, is_batch=is_batch, total_books=len(folders), lang=self.current_lang)
        
        # 创建覆盖策略管理器
        overwrite_policy = OverwritePolicy()
        overwrite_policy.parent_window = self.root

        def finish_callback(success, generated=None, error=None, cancelled=False, lang=None):
            self.root.after(0, lambda: self._on_finish(success, generated, error, cancelled, lang))

        thread = threading.Thread(
            target=run_batch_conversion,
            args=(folders, self.progress_win, finish_callback, self.current_lang, output_dir, overwrite_policy),
            daemon=True
        )
        thread.start()

    def start_merged_conversion(self, base_folder):
        """启动合并转换"""
        self.progress_win = ProgressWindow(self.root, is_batch=False, total_books=1, lang=self.current_lang)
        
        # 创建覆盖策略管理器
        overwrite_policy = OverwritePolicy()
        overwrite_policy.parent_window = self.root

        def finish_callback(success, generated=None, error=None, cancelled=False, lang=None):
            self.root.after(0, lambda: self._on_finish(success, generated, error, cancelled, lang))

        thread = threading.Thread(
            target=run_merged_batch_conversion,
            args=(base_folder, self.progress_win, finish_callback, self.current_lang, overwrite_policy),
            daemon=True
        )
        thread.start()

    def _on_finish(self, success, generated=None, error=None, cancelled=False, lang=None):
        self.progress_win.close()
        tr = self.tr

        if cancelled:
            # 取消时显示已完成的数量
            if generated:
                messagebox.showinfo("操作已取消", f"已成功生成 {len(generated)} 个EPUB文件")
            else:
                messagebox.showinfo("操作已取消", "没有生成任何EPUB文件")
        elif success:
            if generated:
                success_msg = "\n".join([tr("success_epub_created").format(output=epub) for epub in generated])
                messagebox.showinfo("Success", success_msg)
            else:
                messagebox.showinfo("Info", "没有生成任何EPUB文件（所有文件都已跳过）")
        else:
            if error:
                messagebox.showerror("Error", tr("error_conversion_failed").format(error=error))


if __name__ == "__main__":
    check_and_install_deps()
    root = tk.Tk()
    app = App(root)
    root.mainloop()