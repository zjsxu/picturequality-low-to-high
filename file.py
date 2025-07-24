import os
import shutil
import json
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image
import imagehash
import pickle
from concurrent.futures import ThreadPoolExecutor

# ========== 配置项 ==========
HASH_SIZE = 8
SIMILARITY_THRESHOLD = 5  # 阈值：距离 ≤ 5 视为匹配

# ========== UI 更新辅助 ==========
class UiUpdater:
    """将所有对控件的更新调度到主线程执行。"""
    def __init__(self, widget):
        self.widget = widget

    def call(self, fn, *args, **kwargs):
        self.widget.after(0, lambda: fn(*args, **kwargs))

# ========== 主功能函数 ==========
def load_hash_index(index_path, log_widget, progressbar, convert_label):
    log_upd = UiUpdater(log_widget)
    bar_upd = UiUpdater(progressbar)
    lbl_upd = UiUpdater(convert_label)

    pkl_path = index_path + ".pkl"
    log_upd.call(log_widget.insert, tk.END, f"[DEBUG] pkl_path={pkl_path}，存在={os.path.exists(pkl_path)}\n")
    log_upd.call(log_widget.see, tk.END)

    if os.path.exists(pkl_path):
        log_upd.call(log_widget.insert, tk.END, "检测到 .pkl 缓存，开始快速加载索引…\n")
        log_upd.call(log_widget.see, tk.END)
        try:
            with open(pkl_path, 'rb') as f:
                hash_index = pickle.load(f)
        except Exception as e:
            log_upd.call(log_widget.insert, tk.END, f"[ERROR] pickle.load 失败: {e}\n")
            log_upd.call(log_widget.see, tk.END)
            hash_index = {}
        bar_upd.call(progressbar.config, value=100)
        lbl_upd.call(convert_label.config, text=f"已转换 {len(hash_index)}/{len(hash_index)}")
        log_upd.call(log_widget.insert, tk.END, f"快速加载完成，共 {len(hash_index)} 条记录。\n\n")
        log_upd.call(log_widget.see, tk.END)
        return hash_index

    log_upd.call(log_widget.insert, tk.END, "未检测到 .pkl，开始从 JSON 加载索引…\n")
    log_upd.call(log_widget.see, tk.END)
    try:
        with open(index_path, 'r') as jf:
            raw_data = json.load(jf)
            if isinstance(raw_data, list):
                raw = {item['phash']: item['path'] for item in raw_data if 'phash' in item and 'path' in item}
            else:
                raw = raw_data
    except Exception as e:
        log_upd.call(log_widget.insert, tk.END, f"[ERROR] JSON 加载失败: {e}\n")
        log_upd.call(log_widget.see, tk.END)
        return {}

    total = len(raw)

    log_upd.call(log_widget.insert, tk.END, "[DEBUG] JSON 索引前 5 条键：\n")
    for k in list(raw.keys())[:5]:
        log_upd.call(log_widget.insert, tk.END, f"  {k}\n")
    log_upd.call(log_widget.see, tk.END)

    log_upd.call(log_widget.insert, tk.END, f"共 {total} 条哈希记录，逐条转换中：\n")
    log_upd.call(log_widget.see, tk.END)
    lbl_upd.call(convert_label.config, text=f"已转换 0/{total}")

    hash_index = {}

    def convert_item(item):
        hstr, path = item
        try:
            hobj = imagehash.hex_to_hash(hstr)
            return hobj, path
        except Exception as e:
            return None, (hstr, str(e))

    try:
        with ThreadPoolExecutor() as exe:
            for i, result in enumerate(exe.map(convert_item, raw.items()), start=1):
                hobj, path = result
                if hobj is None:
                    log_upd.call(log_widget.insert, tk.END, f"[ERROR] 转换第{i}条失败: {path}\n")
                else:
                    hash_index[hobj] = path

                pct = int(i / total * 100)
                bar_upd.call(progressbar.config, value=pct)
                lbl_upd.call(convert_label.config, text=f"已转换 {i}/{total}")
                if i % max(1, total // 10) == 0 or i == total:
                    log_upd.call(log_widget.insert, tk.END, f"  已转换 {i}/{total} ({pct}%)\n")
                log_upd.call(log_widget.see, tk.END)

    except Exception as e:
        log_upd.call(log_widget.insert, tk.END, f"[ERROR] 转换过程异常: {e}\n")

    if not hash_index:
        log_upd.call(log_widget.insert, tk.END, "[ERROR] 哈希索引为空，检查 JSON 文件内容。\n")
    else:
        log_upd.call(log_widget.insert, tk.END, "[DEBUG] 成功构建哈希索引，前 3 条：\n")
        for h, p in list(hash_index.items())[:3]:
            log_upd.call(log_widget.insert, tk.END, f"  {h} → {p}\n")

    log_upd.call(log_widget.insert, tk.END, "哈希转换完成，开始保存 .pkl 缓存…\n")
    try:
        with open(pkl_path, 'wb') as pf:
            pickle.dump(hash_index, pf)
        log_upd.call(log_widget.insert, tk.END, "[DEBUG] 已成功保存为 .pkl 缓存。\n")
    except Exception as e:
        log_upd.call(log_widget.insert, tk.END, f"[ERROR] 保存 .pkl 失败: {e}\n")
    log_upd.call(log_widget.see, tk.END)

    return hash_index

def find_best_match(low_hash, hash_index, log_widget=None):
    assert isinstance(low_hash, imagehash.ImageHash), f"low_hash 类型错误: {type(low_hash)}"
    
    min_dist, best = float('inf'), None
    for stored_hash, full_path in hash_index.items():
        if not isinstance(stored_hash, imagehash.ImageHash):
            continue  # 跳过非哈希对象
        d = stored_hash - low_hash
        if d < min_dist:
            min_dist, best = d, full_path
            if d == 0:
                break

    if log_widget:
        log_widget.insert(tk.END, f"[DEBUG] 当前比对图与最佳匹配的距离: {min_dist}\n")
        log_widget.see(tk.END)
    print(f"[DEBUG] 当前比对图与最佳匹配的距离: {min_dist}")

    return best if min_dist <= SIMILARITY_THRESHOLD else None


    if log_widget:
        log_widget.insert(tk.END, f"[DEBUG] 当前比对图与最佳匹配的距离: {min_dist}\n")
        log_widget.see(tk.END)
    print(f"[DEBUG] 当前比对图与最佳匹配的距离: {min_dist}")

    return best if min_dist <= SIMILARITY_THRESHOLD else None

def process_images(index_path, input_folder, output_folder, log_widget, progressbar, convert_label):
    log_upd = UiUpdater(log_widget)
    log_widget.insert(tk.END, "步骤 1/2：加载并转换索引\n")
    log_widget.see(tk.END)
    hash_index = load_hash_index(index_path, log_widget, progressbar, convert_label)

    UiUpdater(progressbar).call(progressbar.config, value=0)
    UiUpdater(convert_label).call(convert_label.config, text="")

    log_widget.insert(tk.END, "步骤 2/2：开始图片匹配\n")
    log_widget.see(tk.END)
    unmatched_dir = os.path.join(output_folder, "未找到原图")
    os.makedirs(unmatched_dir, exist_ok=True)
    os.makedirs(output_folder, exist_ok=True)

    files = [f for f in os.listdir(input_folder) if f.lower().endswith(('.jpg','.jpeg','.png'))]
    total = len(files)
    log_widget.insert(tk.END, f"发现 {total} 张图片待匹配\n")
    log_widget.see(tk.END)

    matched, unmatched = 0, 0
    for i, fname in enumerate(files, start=1):
        pct = int(i / total * 100)
        UiUpdater(progressbar).call(progressbar.config, value=pct)
        log_widget.insert(tk.END, f"  处理 {i}/{total}：{fname} … ")
        log_widget.see(tk.END)
        path = os.path.join(input_folder, fname)
        try:
            img = Image.open(path)
            img = img.convert('RGB')  # 防止透明图像影响哈希
            width, height = img.size
            log_widget.insert(tk.END, f"[DEBUG] 图像尺寸: {width}x{height}\n")
            log_widget.see(tk.END)

            low_hash = imagehash.phash(img, hash_size=HASH_SIZE)
            log_widget.insert(tk.END, f"[DEBUG] 当前图片哈希: {str(low_hash)}\n")
            log_widget.see(tk.END)

            match = find_best_match(low_hash, hash_index, log_widget)
            if match:
                shutil.copy2(match, os.path.join(output_folder, fname))
                matched += 1
                log_widget.insert(tk.END, "匹配成功\n")
            else:
                shutil.copy2(path, os.path.join(unmatched_dir, fname))
                unmatched += 1
                log_widget.insert(tk.END, "未匹配，已复制低清图\n")
        except Exception as e:
            log_widget.insert(tk.END, f"[ERROR] 处理失败：{e}\n")
            unmatched += 1
        log_widget.see(tk.END)

    log_widget.insert(tk.END, f"\n完成：匹配 {matched} 张，未匹配 {unmatched} 张。\n")
    log_widget.see(tk.END)
    try:
        with open(os.path.join(output_folder, "整理日志.txt"), 'w', encoding='utf-8') as lf:
            lf.write(f"匹配: {matched}\n未匹: {unmatched}\n")
        log_widget.insert(tk.END, "整理日志已保存。\n")
    except Exception as e:
        log_upd.call(log_widget.insert, tk.END, f"[ERROR] 日志保存失败：{e}\n")
    log_widget.see(tk.END)

    UiUpdater(log_widget).call(lambda: messagebox.showinfo("完成", f"共 {len(files)} 张，匹配 {matched} 张，未匹配 {unmatched} 张。"))

# ========== GUI ==========
class App:
    def __init__(self, root):
        self.root = root
        root.title("图片高清原图匹配工具")
        root.geometry("720x600")

        self.index_path = tk.StringVar()
        self.input_folder = tk.StringVar()
        self.output_folder = tk.StringVar()

        frame = ttk.Frame(root, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="原图哈希索引 JSON：").pack(anchor="w")
        ttk.Entry(frame, textvariable=self.index_path, width=80).pack()
        ttk.Button(frame, text="选择索引文件", command=self.select_index).pack(pady=(0,10))

        ttk.Label(frame, text="低清图文件夹：").pack(anchor="w")
        ttk.Entry(frame, textvariable=self.input_folder, width=80).pack()
        ttk.Button(frame, text="选择低清图文件夹", command=self.select_input).pack(pady=(0,10))

        ttk.Label(frame, text="输出文件夹：").pack(anchor="w")
        ttk.Entry(frame, textvariable=self.output_folder, width=80).pack()
        ttk.Button(frame, text="选择输出文件夹", command=self.select_output).pack(pady=(0,10))

        self.convert_label = ttk.Label(frame, text="已转换 0/0")
        self.convert_label.pack(anchor="w", pady=(5,0))

        ttk.Label(frame, text="运行日志：").pack(anchor="w", pady=(10,0))
        self.log_widget = tk.Text(frame, height=12)
        self.log_widget.pack(fill=tk.BOTH, expand=True)

        self.progressbar = ttk.Progressbar(frame, orient="horizontal", length=600, mode="determinate")
        self.progressbar.pack(pady=(5, 10))

        ttk.Button(frame, text="开始匹配", command=self.start_process).pack()

    def select_index(self):
        path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
        if path:
            self.index_path.set(path)

    def select_input(self):
        path = filedialog.askdirectory()
        if path:
            self.input_folder.set(path)

    def select_output(self):
        path = filedialog.askdirectory()
        if path:
            self.output_folder.set(path)

    def start_process(self):
        if not all([self.index_path.get(), self.input_folder.get(), self.output_folder.get()]):
            messagebox.showerror("错误", "请填写所有路径。")
            return
        self.log_widget.delete("1.0", tk.END)
        self.progressbar["value"] = 0
        self.convert_label.config(text="已转换 0/0")
        threading.Thread(
            target=process_images,
            args=(
                self.index_path.get(),
                self.input_folder.get(),
                self.output_folder.get(),
                self.log_widget,
                self.progressbar,
                self.convert_label
            ),
            daemon=True
        ).start()

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
