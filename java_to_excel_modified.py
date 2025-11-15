import os
import javalang
import pandas as pd
from openpyxl import load_workbook

def parse_java_file(file_path):
    """解析单个Java文件，提取字段信息"""
    with open(file_path, 'r', encoding='utf-8') as f:
        code = f.read()
    
    tree = javalang.parse.parse(code)
    fields = []
    
    for path, node in tree.filter(javalang.tree.FieldDeclaration):
        for declarator in node.declarators:
            # 提取注释
            comment = None
            if node.documentation:
                comment = node.documentation.strip()
            
            # 提取字段名、类型和默认值
            field_info = {
                '字段名': declarator.name,
                '类型': node.type.name,
                '默认值': str(declarator.initializer) if declarator.initializer else None,
                '注释': comment
            }
            fields.append(field_info)
    
    return fields

def process_java_folder(folder_path, output_excel):
    """处理文件夹下所有Java文件，生成多工作表Excel"""
    # 创建ExcelWriter对象
    with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
        # 遍历文件夹中的所有Java文件
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                if file.endswith('.java'):
                    java_file_path = os.path.join(root, file)
                    sheet_name = os.path.splitext(file)[0]  # 工作表名使用文件名（无扩展名）
                    
                    # 解析Java文件
                    fields = parse_java_file(java_file_path)
                    if not fields:
                        continue  # 跳过无字段的文件
                    
                    # 转换为DataFrame并写入工作表
                    df = pd.DataFrame(fields)
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
    
    print(f"生成Excel文件: {output_excel}")

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description='Java字段提取并生成多工作表Excel')
    parser.add_argument('--input_folder', required=True, help='Java文件所在文件夹')
    parser.add_argument('--output_file', required=True, help='输出的Excel文件名')
    args = parser.parse_args()
    
    process_java_folder(args.input_folder, args.output_file)