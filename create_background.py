from PIL import Image, ImageDraw, ImageFont

# 创建一个简单的背景图片
width, height = 800, 600
background_color = (240, 248, 255)  # 淡蓝色背景

# 创建图片
img = Image.new('RGB', (width, height), background_color)
draw = ImageDraw.Draw(img)

# 绘制边框
border_color = (70, 130, 180)  # 钢蓝色
border_width = 5
draw.rectangle([border_width//2, border_width//2, width-border_width//2, height-border_width//2], 
               outline=border_color, width=border_width)

# 添加标题区域
title_area_color = (176, 196, 222)  # 浅钢蓝色
draw.rectangle([20, 20, width-20, 80], fill=title_area_color, outline=border_color, width=2)

# 尝试加载字体，如果失败则使用默认字体
try:
    title_font = ImageFont.truetype("arial.ttf", 24)
    text_font = ImageFont.truetype("arial.ttf", 16)
except:
    title_font = ImageFont.load_default()
    text_font = ImageFont.load_default()

# 添加标题文字
draw.text((width//2-80, 40), "学生信息卡", fill=(25, 25, 112), font=title_font)

# 添加字段标签
labels = ["姓名:", "学院:", "班级:", "辅导员:", "宿舍:", "床位:"]
for i, label in enumerate(labels):
    y_pos = 120 + i * 30
    draw.text((30, y_pos), label, fill=(25, 25, 112), font=text_font)

# 添加照片区域框
photo_x, photo_y = 50, 300
photo_w, photo_h = 150, 150
draw.rectangle([photo_x-2, photo_y-2, photo_x+photo_w+2, photo_y+photo_h+2], 
               outline=border_color, width=2)
draw.text((photo_x+50, photo_y+photo_h+10), "照片", fill=(25, 25, 112), font=text_font)

# 保存图片
img.save('background.jpg', 'JPEG', quality=95)
print("已创建背景图片: background.jpg")
print(f"图片尺寸: {width}x{height}")