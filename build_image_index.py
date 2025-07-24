import os
import json
from PIL import Image
import imagehash
from tqdm import tqdm

def build_image_index(base_dir, output_json="image_index.json"):
    index = []
    valid_exts = (".jpg", ".jpeg", ".png")

    # 收集所有待处理图片路径
    all_images = []
    for root, dirs, files in os.walk(base_dir):
        if "外宣" in root:
            continue
        for file in files:
            if file.lower().endswith(valid_exts):
                all_images.append(os.path.join(root, file))

    print(f"🔍 共发现 {len(all_images)} 张候选图片，开始计算哈希...")

    for path in tqdm(all_images, desc="正在处理图像"):
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
        except:
            continue

    with open(output_json, "w", encoding="utf-8") as f:
        json.dump(index, f, indent=2, ensure_ascii=False)

    print(f"\n✅ 索引构建完成，共索引图像：{len(index)} 张")
    print(f"📝 已保存至：{output_json}")

# 示例用法
if __name__ == "__main__":
    build_image_index("/Volumes/My Passport", "原图索引.json")
