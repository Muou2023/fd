# -*- coding: utf-8 -*-

"""

根据模板文件生成对应的BU罚单
目前需要注意的筛选的条件给架构出勤率不达标给与剔除

"""
import os
import pandas as pd
import shutil
import logging

# 配置日志记录
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# 文件路径
downloadPath = r"D:\Muou\Personal\Downloads"  # 默认下载路径
wpspath = r"D:\Muou\WPSDrive\18133944\WPS云盘\个人文件\质检组文件\田继明"  # 云盘总路径

# 文件路径及文件名称拼接
template_path = os.path.join(wpspath, "质控部重点罚单申请_数据模板.xlsx")  # 模板文件路径
z_fd = os.path.join(wpspath, "日维度扣罚明细.xlsx")  # 日纬度罚单明细表路径

# 检查模板文件或日纬度扣罚明细文件是否存在

if not os.path.isfile(template_path):
    logging.error(f"模板文件 {template_path} 不存在，请检查文件名称和路径是否正确。")
    exit(1)
if not os.path.isfile(z_fd):
    logging.error(f"日纬度扣罚明细 {z_fd} 不存在，请检查文件名称和路径是否正确。")
    exit(1)


# 检查文件路径和目录是否存在
def check_paths(downloadPath, wpspath, z_fd):
    if not os.path.exists(downloadPath):
        logging.error(f"文件路径 {downloadPath} 不存在，请检查路径是否正确。")
        return False
    if not os.path.exists(wpspath):
        logging.error(f"文件路径 {wpspath} 不存在，请检查路径是否正确。")
        return False
    if not os.path.isfile(z_fd):
        logging.error(f"文件 {z_fd} 不存在，请检查文件名称和路径是否正确。")
        return False
    return True


if not check_paths(downloadPath, wpspath, z_fd):
    exit(1)

# 读取总表
data = pd.read_excel(z_fd, sheet_name='日纬度扣罚明细')
logging.info(f"读取的列名: {', '.join(data.columns)}")

# 检查 DataFrame 中是否有必需的列
required_columns = ['BU', '罚单状态(实例ID)', '*违规内容']
if not all(col in data.columns for col in required_columns):
    missing_columns = [col for col in required_columns if col not in data.columns]
    logging.error(f"DataFrame 中没有名为 {'、'.join(missing_columns)} 的列，现有列名: {', '.join(data.columns)}")
    exit(1)

# 筛选罚单状态为空的数据
# filtered_data = data[(data['罚单状态(实例ID)'].isna())]

# 筛选罚单状态为空且违规内容不等于架构出勤率不达标的数据
filtered_data = data[(data['罚单状态(实例ID)'].isna()) & (data['*违规内容'] != '架构出勤率不达标')]
logging.info(f"筛选后的数据数量: {len(filtered_data)}")

# 定义分类映射
categories = {
    '安配': 'M-质控部重点罚单申请_数据模板.xlsx',
    '万物': 'Q-质控部重点罚单申请_数据模板.xlsx',
    '象达': 'X-质控部重点罚单申请_数据模板.xlsx'
}

# 目标表头顺序
target_columns = [
    "*责任人", "*职位", "*所属站点名", "*站点ID", "*所属城市", "*违规类别", "*违规内容", "*处罚规则",
    "*处罚金额", "*罚单产生日期", "*是否连带", "备注", "罚单状态"
]

# 读取模板文件的工作表名称
try:
    template_excel = pd.ExcelFile(template_path)
    sheet_names = template_excel.sheet_names
    logging.info(f"模板文件 {template_path} 中的工作表: {', '.join(sheet_names)}")
    if 'sheet1' not in sheet_names:
        logging.error(f"模板文件 {template_path} 中没有名为 'sheet1' 的工作表，现有工作表: {', '.join(sheet_names)}")
        exit(1)
except Exception as e:
    logging.error(f"读取模板文件 {template_path} 时发生错误: {e}")
    exit(1)

# 为每个分类生成对应的分类表并填充数据
for category, file_name in categories.items():
    # 生成分类表的文件路径
    output_path = os.path.join(downloadPath, file_name)

    # 复制模板文件到分类表路径
    if not os.path.exists(output_path):
        shutil.copy(template_path, output_path)
        logging.info(f"已创建分类表 {output_path}")

    # 读取模板文件
    try:
        template_df = pd.read_excel(output_path, sheet_name='sheet1')
    except Exception as e:
        logging.error(f"读取文件 {output_path} 时发生错误: {e}")
        continue

    # 筛选数据
    filtered_df = filtered_data[filtered_data['BU'].str.contains(category, na=False)]
    logging.info(f"{category} 分类的筛选后的数据数量: {len(filtered_df)}")

    if not filtered_df.empty:
        # 移除 'BU' 列
        filtered_df = filtered_df.drop(columns=['BU'])

        # 确保罚单金额列为数字格式
        filtered_df['*处罚金额'] = pd.to_numeric(filtered_df['*处罚金额'], errors='coerce')
        filtered_df = filtered_df.dropna(subset=['*处罚金额'])

        # 检查并添加缺失的列（默认值为空）
        missing_columns = set(target_columns) - set(filtered_df.columns)
        for col in missing_columns:
            filtered_df[col] = ''

        # 按照目标表头顺序重新排列列
        filtered_df = filtered_df[target_columns]

        # 添加罚单状态列，默认值为“已生效”
        filtered_df['罚单状态'] = '已生效'

        # 将筛选后的数据追加到模板文件中
        try:
            with pd.ExcelWriter(output_path, mode='a', if_sheet_exists='overlay', engine='openpyxl') as writer:
                existing_data = pd.read_excel(output_path, sheet_name='sheet1')
                combined_data = pd.concat([existing_data, filtered_df], ignore_index=True)
                combined_data.to_excel(writer, sheet_name='sheet1', index=False)
            logging.info(f"文件 {output_path} 中的 {category} 分类数据已更新。")
        except Exception as e:
            logging.error(f"更新文件 {output_path} 时发生错误: {e}")
    else:
        logging.info(f"没有找到 {category} 分类的数据。")

logging.info("所有文件的数据填充和更新已完成。")
