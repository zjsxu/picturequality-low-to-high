import os
import shutil
import zipfile
import rarfile
import fitz  # PyMuPDF
import docx
from pathlib import Path
from PIL import Image
from io import BytesIO
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image
import imagehash


# ========== 配置 ==========
ALLOWED_IMG_EXT = {'.jpg', '.jpeg', '.png', '.bmp', '.webp'}
ALLOWED_DOCX = {'.docx'}
ALLOWED_PDF = {'.pdf'}
ALLOWED_ZIP = {'.zip'}
ALLOWED_RAR = {'.rar'}

# ========== 工具函数 ==========

def get_target_year(path):
    for part in path.parts:
        if any(keyword in part for keyword in ["年报", "年", "年年报"]) and part[:4].isdigit():
            return f"{part[:4]}年年报"
    return None

def safe_copy(src, dst_dir):
    dst_dir.mkdir(parents=True, exist_ok=True)
    base_name = Path(src).name
    dst_path = dst_dir / base_name
    counter = 1
    while dst_path.exists():
        dst_path = dst_dir / f"{dst_path.stem}_{counter}{dst_path.suffix}"
        counter += 1
    shutil.copy2(src, dst_path)
    return dst_path.name

def extract_images_from_docx(docx_path, out_dir):
    doc = docx.Document(docx_path)
    rels = doc.part._rels
    count = 0
    for rel in rels:
        rel = rels[rel]
        if "image" in rel.target_ref:
            img_data = rel.target_part.blob
            img = Image.open(BytesIO(img_data))
            out_path = out_dir / f"img_from_docx_{count}.jpg"
            img.save(out_path)
            count += 1
    return count

def extract_images_from_pdf(pdf_path, out_dir):
    doc = fitz.open(pdf_path)
    count = 0
    for page_index in range(len(doc)):
        page = doc[page_index]
        images = page.get_images(full=True)
        for img_index, img in enumerate(images):
            xref = img[0]
            base_image = doc.extract_image(xref)
            img_bytes = base_image["image"]
            img_ext = base_image["ext"]
            img = Image.open(BytesIO(img_bytes))
            out_path = out_dir / f"img_from_pdf_{page_index}_{img_index}.{img_ext}"
            img.save(out_path)
            count += 1
    return count

def extract_from_zip(zip_path, out_dir):
    count = 0
    with zipfile.ZipFile(zip_path, 'r') as zf:
        for f in zf.namelist():
            if Path(f).suffix.lower() in ALLOWED_IMG_EXT:
                target_path = out_dir / Path(f).name
                target_path.parent.mkdir(parents=True, exist_ok=True)
                with zf.open(f) as source, open(target_path, 'wb') as target:
                    shutil.copyfileobj(source, target)
                count += 1
    return count

def extract_from_rar(rar_path, out_dir):
    count = 0
    with rarfile.RarFile(rar_path) as rf:
        for f in rf.namelist():
            if Path(f).suffix.lower() in ALLOWED_IMG_EXT:
                rf.extract(f, out_dir)
                count += 1
    return count

def log_operation(log_path, content, file_path=None):
    with open(log_path, 'a', encoding='utf-8') as f:
        if file_path:
            file_stat = file_path.stat()
            content += f" | 创建时间: {file_stat.st_ctime} | 修改时间: {file_stat.st_mtime}"
        f.write(content + '\n')

def extract_images_from_docx_and_pdf(source_file, target_folder, status_callback):
    """从 docx 或 pdf 文件中提取图片并保存到目标文件夹的子文件夹"""
    source_file = Path(source_file)
    target_folder = Path(target_folder)

    # 创建与文档文件名相同的子文件夹
    sub_folder = target_folder / source_file.stem
    sub_folder.mkdir(parents=True, exist_ok=True)

    try:
        if source_file.suffix.lower() == '.docx':
            # 处理 docx 文件
            doc = docx.Document(source_file)
            rels = doc.part._rels
            count = 0
            for rel in rels:
                rel = rels[rel]
                if "image" in rel.target_ref:
                    img_data = rel.target_part.blob
                    img = Image.open(BytesIO(img_data))
                    out_path = sub_folder / f"img_from_docx_{count}.jpg"
                    img.save(out_path, quality=95)  # 提高保存质量
                    count += 1
            status_callback(f"[DOCX] 提取了 {count} 张图片 <- {source_file}")

        elif source_file.suffix.lower() == '.pdf':
            # 处理 pdf 文件
            doc = fitz.open(source_file)
            count = 0
            for page_index in range(len(doc)):
                page = doc[page_index]
                images = page.get_images(full=True)
                for img_index, img in enumerate(images):
                    xref = img[0]
                    base_image = doc.extract_image(xref)
                    img_bytes = base_image["image"]
                    img_ext = base_image["ext"]
                    img = Image.open(BytesIO(img_bytes))
                    out_path = sub_folder / f"img_from_pdf_{page_index}_{img_index}.{img_ext}"
                    img.save(out_path, quality=95)  # 提高保存质量
                    count += 1
            status_callback(f"[PDF] 提取了 {count} 张图片 <- {source_file}")

        else:
            status_callback(f"[ERROR] 不支持的文件类型: {source_file.suffix}")

    except Exception as e:
        status_callback(f"[ERROR] 提取图片失败: {str(e)}")

def is_similar(img1_path, img2_path, threshold=5):
    try:
        img1 = Image.open(img1_path).convert('RGB')
        img2 = Image.open(img2_path).convert('RGB')
        hash1 = imagehash.phash(img1)
        hash2 = imagehash.phash(img2)
        return hash1 - hash2 <= threshold
    except Exception as e:
        return False

def find_hd_images(low_res_dir, output_dir, index_json, logger=None, threshold=8):
    import json
    import imagehash
    from PIL import Image

    with open(index_json, "r", encoding="utf-8") as f:
        index = json.load(f)

    log = []
    not_found_dir = os.path.join(output_dir, "未找到原图")
    os.makedirs(not_found_dir, exist_ok=True)

    low_images = [f for f in os.listdir(low_res_dir) if f.lower().endswith((".jpg", ".jpeg", ".png"))]

    for img_name in low_images:
        low_img_path = os.path.join(low_res_dir, img_name)
        try:
            low_img = Image.open(low_img_path).convert("RGB")
            low_hash = imagehash.phash(low_img)
        except Exception as e:
            msg = f"❌ 无法读取低清图 {img_name}: {e}"
            log.append(msg)
            if logger: logger(msg)
            continue

        best_match_path = None
        best_match_size = 0

        if logger: logger(f"🔍 正在比对：{img_name}")

        for entry in index:
            try:
                entry_hash = imagehash.hex_to_hash(entry["phash"])
                distance = low_hash - entry_hash
                if distance <= threshold:
                    size = entry["size"][0] * entry["size"][1]
                    if size > best_match_size:
                        best_match_path = entry["path"]
                        best_match_size = size
            except:
                continue

        if best_match_path:
            shutil.copy2(best_match_path, os.path.join(output_dir, img_name))
            msg = f"{img_name} ✅ 匹配成功：{best_match_path}"
        else:
            shutil.copy2(low_img_path, os.path.join(not_found_dir, img_name))
            msg = f"{img_name} ❌ 未找到匹配原图"

        log.append(msg)
        if logger: logger(msg)

    # 写日志
    log_path = os.path.join(output_dir, "matching_log.txt")
    with open(log_path, 'w', encoding='utf-8') as f:
        for line in log:
            f.write(line + "\n")




# ========== 主处理函数 ==========

ALLOWED_IMG_EXT = {'.jpg', '.jpeg', '.png', '.bmp', '.webp'}
ALLOWED_COMPRESSED = {'.zip', '.rar'}

import re

def process_photos(source_root, target_root, status_callback):
    source_root = Path(source_root)
    target_root = Path(target_root)
    total_files = 0

    for root, dirs, files in os.walk(source_root):
        root_path = Path(root)
        year = get_target_year(root_path)
        if not year:
            continue  # 如果无法识别年份，跳过该文件夹

        # 创建目标年份文件夹
        target_year_dir = target_root / year
        target_year_dir.mkdir(parents=True, exist_ok=True)
        log_path = target_year_dir / "log.txt"

        # 用于编号无中文名文件
        no_chinese_counter = 1

        for file in files:
            file_path = root_path / file
            suffix = file_path.suffix.lower()

            try:
                # 如果是图片文件，复制到目标文件夹
                if suffix in ALLOWED_IMG_EXT:
                    # 判断文件名是否包含中文
                    if not contains_chinese(file):
                        # 获取上一级文件夹名称作为事件名称
                        event_name = root_path.name
                        # 生成新的文件名
                        new_file_name = f"{event_name}_{no_chinese_counter}{suffix}"
                        no_chinese_counter += 1
                    else:
                        new_file_name = file

                    # 复制文件并重命名
                    copied_name = safe_copy(file_path, target_year_dir / new_file_name)
                    log_operation(log_path, f"[IMG] {copied_name} <- {file_path}", file_path)
                    status_callback(f"[IMG] {copied_name}")

                # 如果是压缩包，提取图片到目标文件夹
                elif suffix in ALLOWED_COMPRESSED:
                    if suffix == '.zip':
                        count = extract_from_zip(file_path, target_year_dir)
                    elif suffix == '.rar':
                        count = extract_from_rar(file_path, target_year_dir)
                    log_operation(log_path, f"[COMPRESSED] Extracted {count} images <- {file_path}", file_path)
                    status_callback(f"[COMPRESSED] {count} images")

            except Exception as e:
                log_operation(log_path, f"[ERROR] {file_path}: {str(e)}")
                status_callback(f"[ERROR] {file_path.name}")
            total_files += 1

    status_callback(f"处理完成，共处理 {total_files} 个文件")

def contains_chinese(text):
    """判断字符串是否包含中文字符"""
    for char in text:
        if '\u4e00' <= char <= '\u9fff':  # 检查字符是否在中文范围内
            return True
    return False

def build_image_index(base_dir, output_json, logger=None):
    import imagehash
    from PIL import Image
    import json
    import os

    index = []
    valid_exts = (".jpg", ".jpeg", ".png")

    all_images = []
    for root, dirs, files in os.walk(base_dir):
        if "外宣" in root:
            continue
        for file in files:
            if file.lower().endswith(valid_exts):
                all_images.append(os.path.join(root, file))

    if logger: logger(f"📦 共发现图像 {len(all_images)} 张，开始处理...")

    for i, path in enumerate(all_images):
        try:
            with Image.open(path) as img:
                img = img.convert("RGB")
                phash = str(imagehash.phash(img))
                size = img.size
            index.append({
                "path": path,
                "phash": phash,
                "size": size
            })
            if logger and i % 50 == 0:
                logger(f"已完成 {i}/{len(all_images)} 张")
        except Exception as e:
            continue

    with open(output_json, "w", encoding="utf-8") as f:
        json.dump(index, f, indent=2, ensure_ascii=False)

    if logger:
        logger(f"✅ 共索引 {len(index)} 张图像，索引文件已保存")

# ========== GUI ==========

def start_gui():
    import threading

    def browse_source():
        path = filedialog.askdirectory()
        if path:
            src_entry.delete(0, tk.END)
            src_entry.insert(0, path)

    def browse_target():
        path = filedialog.askdirectory()
        if path:
            tgt_entry.delete(0, tk.END)
            tgt_entry.insert(0, path)

    def browse_file():
        path = filedialog.askopenfilename(filetypes=[("文档文件", "*.docx *.pdf")])
        if path:
            file_entry.delete(0, tk.END)
            file_entry.insert(0, path)

    def browse_lowres():
        path = filedialog.askdirectory()
        if path:
            lowres_entry.delete(0, tk.END)
            lowres_entry.insert(0, path)

    def browse_hd_output():
        path = filedialog.askdirectory()
        if path:
            hd_output_entry.delete(0, tk.END)
            hd_output_entry.insert(0, path)

    def browse_hd_search_base():
        path = filedialog.askdirectory()
        if path:
            hd_search_entry.delete(0, tk.END)
            hd_search_entry.insert(0, path)

    def run_photos():
        src = src_entry.get()
        tgt = tgt_entry.get()
        if not src or not tgt:
            messagebox.showerror("错误", "请指定源目录和目标目录")
            return
        status_list.delete(0, tk.END)
        process_photos(src, tgt, lambda msg: status_list.insert(tk.END, msg))

    def run_extract():
        file_path = file_entry.get()
        tgt = tgt_entry.get()
        if not file_path or not tgt:
            messagebox.showerror("错误", "请指定文件和目标目录")
            return
        status_list.delete(0, tk.END)
        extract_images_from_docx_and_pdf(file_path, tgt, lambda msg: status_list.insert(tk.END, msg))

    def threaded_hd_match():
        low = lowres_entry.get()
        out = hd_output_entry.get()
        base = hd_search_entry.get()
        if not low or not out or not base:
            messagebox.showerror("错误", "请指定全部三个路径")
            return
        status_list.delete(0, tk.END)
        status_list.insert(tk.END, "🔍 正在查找高清原图，请稍候...")

        def run():
            def log(msg):
                status_list.insert(tk.END, msg)
                status_list.yview_moveto(1)

            find_hd_images(low, out, base, log)
            status_list.insert(tk.END, "✅ 查找完成，结果已输出到目标文件夹")

        threading.Thread(target=run, daemon=True).start()

    root = tk.Tk()
    root.title("年报照片归档助手")

        # ===== 哈希索引构建部分 =====
    tk.Label(root, text="原图根目录（排除外宣）:").grid(row=10, column=0, sticky="e")
    index_src_entry = tk.Entry(root, width=60)
    index_src_entry.grid(row=10, column=1)
    tk.Button(root, text="浏览", command=lambda: index_src_entry.insert(0, filedialog.askdirectory())).grid(row=10, column=2)

    tk.Label(root, text="索引保存为:").grid(row=11, column=0, sticky="e")
    index_out_entry = tk.Entry(root, width=60)
    index_out_entry.grid(row=11, column=1)
    index_out_entry.insert(0, "原图索引.json")  # 默认值

    def run_build_index():
        import threading
        src_dir = index_src_entry.get()
        out_file = index_out_entry.get()
        if not src_dir or not out_file:
            messagebox.showerror("错误", "请指定源目录和索引保存路径")
            return
        status_list.delete(0, tk.END)
        status_list.insert(tk.END, "🔍 正在构建哈希索引，请稍候...")

        def task():
            build_image_index(src_dir, out_file, lambda msg: status_list.insert(tk.END, msg))
            status_list.insert(tk.END, f"✅ 索引构建完成，已保存至：{out_file}")

        threading.Thread(target=task, daemon=True).start()

    tk.Button(root, text="生成原图哈希索引", command=run_build_index, bg="purple", fg="white").grid(row=12, column=1, pady=10)

    # ===== 整理图片部分 =====
    tk.Label(root, text="源目录（移动硬盘）:").grid(row=0, column=0, sticky="e")
    src_entry = tk.Entry(root, width=60)
    src_entry.grid(row=0, column=1)
    tk.Button(root, text="浏览", command=browse_source).grid(row=0, column=2)

    tk.Label(root, text="目标目录（整理输出）:").grid(row=1, column=0, sticky="e")
    tgt_entry = tk.Entry(root, width=60)
    tgt_entry.grid(row=1, column=1)
    tk.Button(root, text="浏览", command=browse_target).grid(row=1, column=2)

    tk.Button(root, text="整理图片", command=run_photos, bg="green", fg="black").grid(row=2, column=1, pady=10)

    # ===== 提取文档图片部分 =====
    tk.Label(root, text="选择文件（docx/pdf）:").grid(row=3, column=0, sticky="e")
    file_entry = tk.Entry(root, width=60)
    file_entry.grid(row=3, column=1)
    tk.Button(root, text="浏览", command=browse_file).grid(row=3, column=2)

    tk.Button(root, text="提取文档图片", command=run_extract, bg="blue", fg="black").grid(row=4, column=1, pady=10)

    # ===== 查找高清原图部分 =====
    tk.Label(root, text="低清图片文件夹:").grid(row=5, column=0, sticky="e")
    lowres_entry = tk.Entry(root, width=60)
    lowres_entry.grid(row=5, column=1)
    tk.Button(root, text="浏览", command=browse_lowres).grid(row=5, column=2)

    tk.Label(root, text="输出文件夹:").grid(row=6, column=0, sticky="e")
    hd_output_entry = tk.Entry(root, width=60)
    hd_output_entry.grid(row=6, column=1)
    tk.Button(root, text="浏览", command=browse_hd_output).grid(row=6, column=2)

    tk.Label(root, text="查找原图的根目录:").grid(row=7, column=0, sticky="e")
    hd_search_entry = tk.Entry(root, width=60)
    hd_search_entry.grid(row=7, column=1)
    tk.Button(root, text="浏览", command=browse_hd_search_base).grid(row=7, column=2)

    tk.Button(root, text="查找高清原图", command=threaded_hd_match, bg="orange", fg="black").grid(row=8, column=1, pady=10)

    # 状态输出
    status_list = tk.Listbox(root, width=100, height=20)
    status_list.grid(row=9, column=0, columnspan=3, padx=10, pady=10)

    root.mainloop()


if __name__ == "__main__":
    start_gui()
