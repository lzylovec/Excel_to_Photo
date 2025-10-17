import pandas as pd
import os
from PIL import Image, ImageDraw, ImageFont

def create_sample_photos():
    """创建示例照片文件"""
    photos_dir = "photos"
    os.makedirs(photos_dir, exist_ok=True)
    
    # 创建示例照片
    students = [
        {"name": "张三", "color": "#FF6B6B"},
        {"name": "李四", "color": "#4ECDC4"},
        {"name": "王五", "color": "#45B7D1"},
        {"name": "赵六", "color": "#96CEB4"}
    ]
    
    photo_paths = []
    
    for i, student in enumerate(students):
        # 创建一个简单的头像图片
        img = Image.new('RGB', (150, 200), color=student["color"])
        draw = ImageDraw.Draw(img)
        
        # 尝试加载字体
        try:
            font = ImageFont.truetype("/System/Library/Fonts/Arial.ttf", 24)
        except:
            font = ImageFont.load_default()
        
        # 在图片中央绘制姓名
        text = student["name"]
        bbox = draw.textbbox((0, 0), text, font=font)
        text_width = bbox[2] - bbox[0]
        text_height = bbox[3] - bbox[1]
        
        x = (150 - text_width) // 2
        y = (200 - text_height) // 2
        
        draw.text((x, y), text, fill="white", font=font)
        
        # 保存图片
        photo_filename = f"student_{i+1}.jpg"
        photo_path = os.path.join(photos_dir, photo_filename)
        img.save(photo_path, "JPEG", quality=95)
        
        # 使用相对路径
        photo_paths.append(f"photos/{photo_filename}")
        print(f"已创建照片: {photo_path}")
    
    return photo_paths

def create_excel_with_local_photos():
    """创建包含本地图片路径的Excel文件"""
    # 创建示例照片
    photo_paths = create_sample_photos()
    
    # 创建示例数据
    data = {
        '姓名': ['张三', '李四', '王五', '赵六'],
        '学院': ['计算机学院', '电子工程学院', '机械工程学院', '经济管理学院'],
        '班级': ['计算机2021-1班', '电子2021-2班', '机械2021-3班', '经管2021-1班'],
        '辅导员': ['陈老师', '刘老师', '王老师', '李老师'],
        '寝室号': ['A101', 'B202', 'C303', 'D404'],
        '床位号': ['1', '2', '3', '4'],
        '图片 URL': photo_paths
    }
    
    # 创建DataFrame
    df = pd.DataFrame(data)
    
    # 保存为Excel文件
    excel_filename = 'data_local_photos.xlsx'
    df.to_excel(excel_filename, index=False)
    print(f"已创建Excel文件: {excel_filename}")
    print("包含列:", list(df.columns))
    print("本地图片路径:")
    for i, path in enumerate(photo_paths):
        print(f"  {i+1}. {path}")

if __name__ == "__main__":
    create_excel_with_local_photos()