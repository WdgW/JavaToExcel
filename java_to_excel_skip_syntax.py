import os
import re
import javalang
import pandas as pd
import logging
from javalang.parser import JavaSyntaxError

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('parsing_errors.log'),
        logging.StreamHandler()
    ]
)

def preprocess_java_code(code):
    """预处理Java代码：移除注解和泛型语法"""
    # 移除注解（@开头的行）
    code = re.sub(r'@\w+\s+', '', code)
    code = re.sub(r'@\w+\(.*?\)', '', code)
    
    # 移除泛型符号（保留尖括号内的内容但移除尖括号）
    code = re.sub(r'<.*?>', '', code)
    
    # 移除单行注释
    code = re.sub(r'//.*', '', code)
    
    # 移除多行注释
    code = re.sub(r'/\*.*?\*/', '', code, flags=re.DOTALL)
    
    return code

def parse_java_file(file_path):
    """解析预处理后的Java文件"""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            original_code = f.read()
        
        # 预处理代码
        processed_code = preprocess_java_code(original_code)
        
        # 解析处理后的代码
        tree = javalang.parse.parse(processed_code)
        fields = []
        
        for path, node in tree.filter(javalang.tree.FieldDeclaration):
            for declarator in node.declarators:
                field_info = {
                    '字段名': declarator.name,
                    '类型': node.type.name,
                    '默认值': str(declarator.initializer) if declarator.initializer else None,
                    '注释': node.documentation.strip() if node.documentation else None
                }
                fields.append(field_info)
        
        return fields
    
    except JavaSyntaxError as e:
        logging.error(f"Java语法错误 in {file_path}: {str(e)}")
    except Exception as e:
        logging.error(f"解析文件失败 {file_path}: {str(e)}")
    return None

def process_java_folder(folder_path, output_excel):
    """处理文件夹中的Java文件"""
    with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
        success_count = 0
        error_count = 0
        
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                if file.endswith('.java'):
                    java_file_path = os.path.join(root, file)
                    sheet_name = os.path.splitext(file)[0]
                    
                    # 限制工作表名长度
                    if len(sheet_name) > 31:
                        sheet_name = sheet_name[:28] + '...'
                    
                    fields = parse_java_file(java_file_path)
                    if fields:
                        df = pd.DataFrame(fields)
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                        success_count += 1
                    else:
                        error_count += 1
        
        logging.info(f"处理完成: {success_count}个文件成功, {error_count}个文件失败")
        print(f"生成Excel文件: {output_excel}")
        print(f"错误日志已保存至: parsing_errors.log")

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description='Java字段提取（跳过注解和泛型）')
    parser.add_argument('--input_folder', required=True, help='Java文件所在文件夹')
    parser.add_argument('--output_file', required=True, help='输出的Excel文件名')
    args = parser.parse_args()
    
    process_java_folder(args.input_folder, args.output_file)