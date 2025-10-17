from flask import Flask, request, jsonify, send_from_directory, render_template, send_file
from werkzeug.utils import secure_filename
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import requests
from io import BytesIO
import os
import uuid
import traceback
from datetime import datetime
import zipfile

app = Flask(__name__)

# 配置
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'static/output'
ALLOWED_EXCEL_EXTENSIONS = {'xlsx', 'xls'}
ALLOWED_IMAGE_EXTENSIONS = {'png', 'jpg', 'jpeg'}

# 创建必要的目录
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
os.makedirs('static', exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

def allowed_file(filename, allowed_extensions):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in allowed_extensions

@app.route('/')
def index():
    return send_from_directory('.', 'index.html')

@app.route('/<path:filename>')
def serve_static(filename):
    return send_from_directory('.', filename)

@app.route('/generate', methods=['POST'])
def generate_cards():
    try:
        # 检查是否有文件上传
        if 'excel_file' not in request.files or 'image_file' not in request.files:
            return jsonify({'error': '请上传Excel文件和背景图片'}), 400
        
        excel_file = request.files['excel_file']
        background_file = request.files['image_file']
        
        # 检查文件是否为空
        if excel_file.filename == '' or background_file.filename == '':
            return jsonify({'error': '请选择有效的文件'}), 400
        
        # 检查文件类型
        if not (allowed_file(excel_file.filename, ALLOWED_EXCEL_EXTENSIONS) and 
                allowed_file(background_file.filename, ALLOWED_IMAGE_EXTENSIONS)):
            return jsonify({'error': '文件类型不支持'}), 400
        
        # 获取寝室类型参数
        room_type = int(request.form.get('room_type', 4))
        if room_type not in [2, 3, 4, 5]:
            return jsonify({'error': '不支持的寝室类型'}), 400
        
        # 获取位置调整参数
        position_adjustments = {}
        if 'position_adjustments' in request.form:
            import json
            position_adjustments = json.loads(request.form.get('position_adjustments', '{}'))
        
        # 生成唯一的会话ID
        session_id = str(uuid.uuid4())
        session_folder = os.path.join(app.config['OUTPUT_FOLDER'], session_id)
        os.makedirs(session_folder, exist_ok=True)
        
        # 保存上传的文件
        excel_filename = secure_filename(excel_file.filename)
        background_filename = secure_filename(background_file.filename)
        
        excel_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{session_id}_{excel_filename}")
        background_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{session_id}_{background_filename}")
        
        excel_file.save(excel_path)
        background_file.save(background_path)
        
        # 处理图片生成
        result = process_images(
            excel_path, 
            background_path, 
            session_folder, 
            session_id,
            room_type,
            position_adjustments
        )
        
        # 清理临时文件
        os.remove(excel_path)
        os.remove(background_path)
        
        return jsonify(result)
        
    except Exception as e:
        print(f"生成信息卡时发生错误: {str(e)}")
        print(traceback.format_exc())
        return jsonify({'error': f'生成失败: {str(e)}'}), 500

def get_room_layout_positions(room_type, bg_width, bg_height, default_text_x=200, default_text_y=50, default_photo_x=50, default_photo_y=50):
    """
    根据寝室类型计算布局位置
    
    Args:
        room_type: 寝室类型 (2, 3, 4, 5)
        bg_width: 背景图宽度
        bg_height: 背景图高度
        default_text_x: 默认文本X坐标
        default_text_y: 默认文本Y坐标
        default_photo_x: 默认照片X坐标
        default_photo_y: 默认照片Y坐标
    
    Returns:
        位置列表，每个位置包含text_start和photo_pos
    """
    positions = []
    
    if room_type == 2:
        # 双人寝：上下布局
        positions = [
            {'text_start': (default_text_x, default_text_y), 'photo_pos': (default_photo_x, default_photo_y)},  # 上
            {'text_start': (default_text_x, bg_height//2 + default_text_y), 'photo_pos': (default_photo_x, bg_height//2 + default_photo_y)}  # 下
        ]
    elif room_type == 3:
        # 三人寝：上1下2布局
        positions = [
            {'text_start': (bg_width//2 - 100, default_text_y), 'photo_pos': (bg_width//2 - 150, default_photo_y)},  # 上中
            {'text_start': (default_text_x, bg_height//2 + default_text_y), 'photo_pos': (default_photo_x, bg_height//2 + default_photo_y)},  # 下左
            {'text_start': (bg_width//2 + default_text_x, bg_height//2 + default_text_y), 'photo_pos': (bg_width//2 + default_photo_x, bg_height//2 + default_photo_y)}  # 下右
        ]
    elif room_type == 4:
        # 四人寝：2x2网格布局
        positions = [
            {'text_start': (default_text_x, default_text_y), 'photo_pos': (default_photo_x, default_photo_y)},  # 左上
            {'text_start': (bg_width//2 + default_text_x, default_text_y), 'photo_pos': (bg_width//2 + default_photo_x, default_photo_y)},  # 右上
            {'text_start': (default_text_x, bg_height//2 + default_text_y), 'photo_pos': (default_photo_x, bg_height//2 + default_photo_y)},  # 左下
            {'text_start': (bg_width//2 + default_text_x, bg_height//2 + default_text_y), 'photo_pos': (bg_width//2 + default_photo_x, bg_height//2 + default_photo_y)}  # 右下
        ]
    elif room_type == 5:
        # 五人寝：上2下3布局
        positions = [
            {'text_start': (bg_width//4, default_text_y), 'photo_pos': (bg_width//4 - 50, default_photo_y)},  # 上左
            {'text_start': (3*bg_width//4, default_text_y), 'photo_pos': (3*bg_width//4 - 50, default_photo_y)},  # 上右
            {'text_start': (bg_width//6, bg_height//2 + default_text_y), 'photo_pos': (bg_width//6 - 50, bg_height//2 + default_photo_y)},  # 下左
            {'text_start': (bg_width//2, bg_height//2 + default_text_y), 'photo_pos': (bg_width//2 - 50, bg_height//2 + default_photo_y)},  # 下中
            {'text_start': (5*bg_width//6, bg_height//2 + default_text_y), 'photo_pos': (5*bg_width//6 - 50, bg_height//2 + default_photo_y)}  # 下右
        ]
    
    return positions


def process_images(excel_path, background_path, session_folder, session_id, room_type, position_adjustments=None):
    """
    处理图片生成
    
    Args:
        excel_path: Excel文件路径
        background_path: 背景图片路径
        session_folder: 会话文件夹路径
        session_id: 会话ID
        room_type: 寝室类型 (2, 3, 4, 5)
        position_adjustments: 位置调整参数字典
    """
    try:
        # 读取Excel文件
        df = pd.read_excel(excel_path)
        
        if len(df) == 0:
            return {'success': False, 'error': 'Excel文件为空'}
        
        # 检查必要的列
        required_columns = ['姓名', '学院', '班级', '辅导员', '寝室号', '床位号']
        
        # 图片URL列支持多种格式
        photo_url_columns = ['图片 URL', '图片URL', '照片URL', '照片 URL']
        photo_column = None
        
        # 检查基本列
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        # 检查图片URL列
        for col in photo_url_columns:
            if col in df.columns:
                photo_column = col
                break
        
        if photo_column is None:
            missing_columns.append('图片URL列（支持：图片 URL、图片URL、照片URL、照片 URL）')
        
        if missing_columns:
            return {'success': False, 'error': f'Excel文件缺少必要的列: {", ".join(missing_columns)}'}
        
        # 加载背景图
        background = Image.open(background_path).convert("RGB")
        bg_width, bg_height = background.size
        
        # 默认参数
        default_font_size = 20
        default_photo_width = 100
        default_photo_height = 120
        default_photo_x = 50
        default_photo_y = 50
        default_text_x = 200
        default_text_y = 50
        default_line_spacing = 30
        
        # 尝试加载字体 - 优先使用支持中文的字体
        try:
            # 尝试加载macOS系统中文字体
            font = ImageFont.truetype("/System/Library/Fonts/PingFang.ttc", default_font_size)
        except:
            try:
                # 备选中文字体
                font = ImageFont.truetype("/System/Library/Fonts/STHeiti Light.ttc", default_font_size)
            except:
                try:
                    # 再次备选
                    font = ImageFont.truetype("/System/Library/Fonts/Arial Unicode MS.ttf", default_font_size)
                except:
                    try:
                        # 最后尝试Arial
                        font = ImageFont.truetype("/System/Library/Fonts/Arial.ttf", default_font_size)
                    except:
                        # 使用默认字体
                        font = ImageFont.load_default()
        
        generated_images = []
        student_info = []  # 存储学生信息用于前端选择器
        
        # 初始化位置调整
        if position_adjustments is None:
            position_adjustments = {}
        
        # 将学生数据按寝室号进行分组
        students_per_card = room_type
        total_students = len(df)
        
        # 按寝室号分组
        grouped_by_dorm = df.groupby('寝室号')
        
        card_idx = 0
        dorm_info = []  # 记录每个寝室的信息
        for dorm_number, dorm_group in grouped_by_dorm:
            # 获取当前寝室的学生
            current_group = dorm_group
            actual_students_count = len(current_group)
            dorm_info.append({
                'dorm_number': dorm_number,
                'student_count': actual_students_count
            })
            
            # 创建背景副本
            img = background.copy()
            draw = ImageDraw.Draw(img)
            
            # 根据寝室类型计算布局位置
            positions = get_room_layout_positions(room_type, bg_width, bg_height, default_text_x, default_text_y, default_photo_x, default_photo_y)
            
            # 处理当前组的每个学生（只处理实际存在的学生数量）
            for student_idx, (_, row) in enumerate(current_group.iterrows()):
                if student_idx >= room_type:  # 最多room_type个学生
                    break
                
                pos = positions[student_idx]
                position_id = student_idx + 1
                
                # 获取学生数据，处理可能的NaN值
                name = str(row['姓名']) if pd.notna(row['姓名']) else '未填写'
                college = str(row['学院']) if pd.notna(row['学院']) else '未填写'
                class_ = str(row['班级']) if pd.notna(row['班级']) else '未填写'
                counselor = str(row['辅导员']) if pd.notna(row['辅导员']) else '未填写'
                dorm = str(row['寝室号']) if pd.notna(row['寝室号']) else '未填写'
                bed = str(row['床位号']) if pd.notna(row['床位号']) else '未填写'
                photo_url = str(row[photo_column]) if pd.notna(row[photo_column]) else ''
                
                # 添加学生信息到列表
                student_key = f"card_{card_idx}_student_{student_idx}"
                student_info.append({
                    'key': student_key,
                    'name': name,
                    'card_number': card_idx + 1,
                    'position_in_card': student_idx + 1
                })
                
                # 检查是否有该位置的调整参数
                position_key = str(position_id)  # 位置1-room_type
                if position_key in position_adjustments:
                    adjustment = position_adjustments[position_key]
                    # 应用位置调整
                    adjusted_text_pos = (
                        pos['text_start'][0] + adjustment.get('text_x', 0),
                        pos['text_start'][1] + adjustment.get('text_y', 0)
                    )
                    adjusted_photo_pos = (
                        pos['photo_pos'][0] + adjustment.get('photo_x', 0),
                        pos['photo_pos'][1] + adjustment.get('photo_y', 0)
                    )
                    # 应用照片尺寸调整
                    adjusted_photo_width = default_photo_width + adjustment.get('photo_width', 0) if adjustment.get('photo_width', 0) != 0 else default_photo_width
                    adjusted_photo_height = default_photo_height + adjustment.get('photo_height', 0) if adjustment.get('photo_height', 0) != 0 else default_photo_height
                    # 应用字体大小调整
                    adjusted_font_size = default_font_size + adjustment.get('font_size', 0) if adjustment.get('font_size', 0) != 0 else default_font_size
                    # 应用行间距调整
                    adjusted_line_spacing = default_line_spacing + adjustment.get('line_spacing', 0) if adjustment.get('line_spacing', 0) != 0 else default_line_spacing
                    
                    # 创建调整后的字体 - 使用与主字体相同的加载逻辑
                    try:
                        # 尝试加载macOS系统中文字体
                        adjusted_font = ImageFont.truetype("/System/Library/Fonts/PingFang.ttc", adjusted_font_size)
                    except:
                        try:
                            # 备选中文字体
                            adjusted_font = ImageFont.truetype("/System/Library/Fonts/STHeiti Light.ttc", adjusted_font_size)
                        except:
                            try:
                                # 再次备选
                                adjusted_font = ImageFont.truetype("/System/Library/Fonts/Arial Unicode MS.ttf", adjusted_font_size)
                            except:
                                try:
                                    # 最后尝试Arial
                                    adjusted_font = ImageFont.truetype("/System/Library/Fonts/Arial.ttf", adjusted_font_size)
                                except:
                                    # 使用默认字体
                                    adjusted_font = ImageFont.load_default()
                else:
                    adjusted_text_pos = pos['text_start']
                    adjusted_photo_pos = pos['photo_pos']
                    adjusted_photo_width = default_photo_width
                    adjusted_photo_height = default_photo_height
                    adjusted_font_size = default_font_size
                    adjusted_line_spacing = default_line_spacing
                    adjusted_font = font
                
                # 文本信息 - 只显示数据值，不显示标签
                text_items = [
                    name,
                    college,
                    class_,
                    counselor,
                    dorm,
                    bed
                ]
                
                # 绘制文本 - 使用调整后的位置和用户自定义的行间距
                text_x_pos, text_y_pos = adjusted_text_pos
                
                for i, item in enumerate(text_items):
                    y_pos = text_y_pos + i * adjusted_line_spacing
                    # 确保文本不会超出边界
                    if y_pos + adjusted_font_size < bg_height:
                        try:
                            # 确保文本是字符串类型并处理None值
                            text_to_draw = str(item) if item is not None else ""
                            # 绘制文本，使用黑色填充和调整后的字体
                            draw.text((text_x_pos, y_pos), text_to_draw, fill=(0, 0, 0), font=adjusted_font)
                        except Exception as e:
                            print(f"文本绘制错误: {e}, 文本内容: {item}")
                            # 如果绘制失败，尝试处理特殊字符
                            try:
                                # 确保是字符串并移除可能的特殊字符
                                safe_item = str(item) if item is not None else ""
                                safe_item = safe_item.replace('\x00', '').replace('\ufffd', '')  # 移除空字符和替换字符
                                draw.text((text_x_pos, y_pos), safe_item, fill=(0, 0, 0), font=adjusted_font)
                            except:
                                # 最后的备选方案
                                draw.text((text_x_pos, y_pos), "数据显示异常", fill=(0, 0, 0), font=adjusted_font)
                
                # 处理照片
                if photo_url and photo_url.strip():
                    try:
                        if photo_url.startswith(('http://', 'https://')):
                            # 网络图片 - 跳过处理
                            print(f"⚠️  网络图片URL无法访问，跳过: {photo_url}")
                            print(f"   建议：请将图片下载到本地 photos/ 目录，或检查URL是否正确")
                        else:
                            # 本地文件路径
                            photo_path = photo_url
                            if not os.path.isabs(photo_path):
                                # 相对路径，在当前目录下查找
                                photo_path = os.path.join(os.getcwd(), photo_path)
                            
                            if os.path.exists(photo_path):
                                photo = Image.open(photo_path).convert("RGBA")
                                print(f"✅ 成功加载本地图片: {photo_path}")
                                
                                # 调整照片大小 - 使用调整后的尺寸
                                photo = photo.resize((adjusted_photo_width, adjusted_photo_height), Image.Resampling.LANCZOS)
                                
                                # 粘贴照片 - 使用调整后的位置
                                photo_x_pos, photo_y_pos = adjusted_photo_pos
                                img.paste(photo, (photo_x_pos, photo_y_pos), photo if photo.mode == "RGBA" else None)
                            else:
                                print(f"❌ 本地图片文件不存在: {photo_path}")
                        
                    except Exception as e:
                        print(f"❌ 图片加载失败: {photo_url}, 错误: {e}")
                else:
                    print(f"⚠️  跳过空的照片URL")
            
            # 保存卡片
            card_number = card_idx + 1
            output_filename = f"{dorm_number}.jpg"
            output_path = os.path.join(session_folder, output_filename)
            # 保存时不设置quality参数，避免JPEG压缩影响色彩
            img.save(output_path, "JPEG")
            
            # 添加到结果列表
            web_path = f"/static/output/{session_id}/{output_filename}"
            generated_images.append(web_path)
            
            # 增加卡片索引
            card_idx += 1
        
        # 获取学院信息（从第一个学生的数据中获取，假设同一批次都是同一学院）
        college_name = '未知学院'
        if not df.empty and '学院' in df.columns:
            first_college = df['学院'].iloc[0]
            if pd.notna(first_college):
                college_name = str(first_college)
        
        return {
            'success': True,
            'images': generated_images,
            'count': len(generated_images),
            'total_students': total_students,
            'students_per_card': room_type,  # 保持原有的寝室类型信息
            'student_info': student_info,
            'dorm_info': dorm_info,  # 添加寝室信息
            'college_name': college_name  # 添加学院信息用于ZIP命名
        }
        
    except Exception as e:
        print(f"Error in process_images: {e}")
        traceback.print_exc()
        return {'success': False, 'error': f'图片处理错误: {str(e)}'}

@app.route('/static/output/<session_id>/<filename>')
def serve_output_image(session_id, filename):
    return send_from_directory(os.path.join(OUTPUT_FOLDER, session_id), filename)

@app.route('/download_zip/<session_id>')
def download_zip(session_id):
    """下载指定会话的所有图片打包成ZIP文件"""
    try:
        session_folder = os.path.join(OUTPUT_FOLDER, session_id)
        
        # 检查会话文件夹是否存在
        if not os.path.exists(session_folder):
            return jsonify({'success': False, 'error': '会话不存在或已过期'})
        
        # 获取文件夹中的所有图片文件
        image_files = []
        for filename in os.listdir(session_folder):
            if filename.lower().endswith(('.jpg', '.jpeg', '.png')):
                image_files.append(filename)
        
        if not image_files:
            return jsonify({'success': False, 'error': '没有找到可下载的图片文件'})
        
        # 获取前端传递的文件名参数，如果没有则使用默认名称
        custom_filename = request.args.get('filename', f"学生信息卡合集_{datetime.now().strftime('%Y%m%d_%H%M%S')}")
        
        # 创建ZIP文件
        zip_filename = f"{custom_filename}.zip"
        zip_path = os.path.join(session_folder, zip_filename)
        
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for image_file in image_files:
                image_path = os.path.join(session_folder, image_file)
                # 在ZIP中使用更友好的文件名
                arcname = f"学生信息卡_{image_file}"
                zipf.write(image_path, arcname)
        
        # 发送ZIP文件
        return send_file(
            zip_path,
            as_attachment=True,
            download_name=f"{custom_filename}.zip",
            mimetype='application/zip'
        )
        
    except Exception as e:
        print(f"Error in download_zip: {e}")
        traceback.print_exc()
        return jsonify({'success': False, 'error': f'ZIP文件生成错误: {str(e)}'})

@app.errorhandler(413)
def too_large(e):
    return jsonify({'success': False, 'error': '文件太大，请上传小于16MB的文件'})

@app.errorhandler(404)
def not_found(e):
    return jsonify({'success': False, 'error': '请求的资源不存在'})

@app.errorhandler(500)
def server_error(e):
    return jsonify({'success': False, 'error': '服务器内部错误'})

if __name__ == '__main__':
    print("启动学生信息卡生成器服务...")
    print("访问地址: http://localhost:5002")
    app.run(debug=True, host='0.0.0.0', port=5002)