import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import requests
from io import BytesIO
import os

# 配置项
BACKGROUND_IMG_PATH = "background.jpg"  # 用户上传的背景图
EXCEL_PATH = "data.xlsx"                # 用户上传的Excel
OUTPUT_DIR = "output"                   # 输出目录
TEXT_POSITIONS = [(50, 50), (50, 80), (50, 110), (50, 140), (50, 170), (50, 200)]  # 文本位置列表
PHOTO_POSITION = (50, 300)              # 照片粘贴位置
PHOTO_SIZE = (150, 150)                 # 照片缩放尺寸
FONT_PATH = "arial.ttf"                 # 字体文件（可选，若无则用默认）
FONT_SIZE = 20

# 创建输出目录
os.makedirs(OUTPUT_DIR, exist_ok=True)

# 加载背景图（作为模板）
background = Image.open(BACKGROUND_IMG_PATH).convert("RGB")

# 读取Excel
df = pd.read_excel(EXCEL_PATH)

# 尝试加载字体
try:
    font = ImageFont.truetype(FONT_PATH, FONT_SIZE)
except:
    font = ImageFont.load_default()

# 遍历每一行
for idx, row in df.iterrows():
    # 创建背景副本
    img = background.copy()
    draw = ImageDraw.Draw(img)

    # 获取文本信息和照片URL
    name, college, class_, counselor, dorm, bed, photo_url = row
    
    text_items = [name, college, class_, counselor, dorm, bed]

    # 绘制文本
    for i, item in enumerate(text_items):
        draw.text(TEXT_POSITIONS[i], str(item), fill="black", font=font)

    # 加载并粘贴照片
    try:
        if photo_url.lower().startswith('http'):
            response = requests.get(photo_url)
            photo = Image.open(BytesIO(response.content)).convert("RGBA")
        else:
            photo = Image.open(photo_url).convert("RGBA")
        
        photo = photo.resize(PHOTO_SIZE, Image.Resampling.LANCZOS)
        img.paste(photo, PHOTO_POSITION, photo if photo.mode == "RGBA" else None)
    except Exception as e:
        print(f"警告：无法加载照片 {photo_url}，跳过。错误：{e}")

    # 保存结果
    output_path = os.path.join(OUTPUT_DIR, f"output_{idx}.jpg")
    img.save(output_path, "JPEG")
    print(f"已保存：{output_path}")