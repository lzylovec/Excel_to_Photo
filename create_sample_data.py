import pandas as pd

# 创建示例数据
data = {
    '姓名': ['张三', '李四', '王五', '赵六'],
    '学院': ['计算机学院', '电子工程学院', '机械工程学院', '经济管理学院'],
    '班级': ['计算机2021-1班', '电子2021-2班', '机械2021-3班', '经管2021-1班'],
    '辅导员': ['陈老师', '刘老师', '王老师', '李老师'],
    '宿舍': ['1号楼', '2号楼', '3号楼', '4号楼'],
    '床位': ['101-1', '202-2', '303-3', '404-4'],
    '照片URL': [
        'https://via.placeholder.com/150x150/FF0000/FFFFFF?text=张三',
        'https://via.placeholder.com/150x150/00FF00/FFFFFF?text=李四',
        'https://via.placeholder.com/150x150/0000FF/FFFFFF?text=王五',
        'https://via.placeholder.com/150x150/FFFF00/000000?text=赵六'
    ]
}

# 创建DataFrame并保存为Excel
df = pd.DataFrame(data)
df.to_excel('data.xlsx', index=False)
print("已创建示例Excel文件: data.xlsx")
print("数据内容:")
print(df)