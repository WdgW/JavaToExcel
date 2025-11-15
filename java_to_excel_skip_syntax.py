import os
import re
import javalang
import pandas as pd
import logging
from javalang.parser import JavaSyntaxError
from javalang.ast import Node

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('parsing_errors.log'),
        logging.StreamHandler()
    ]
)

def get_literal_value(node):
    """提取Literal节点的值"""
    if isinstance(node, javalang.tree.Literal):
        return node.value
    return None

def get_node_text(node, code):
    """获取AST节点在原始代码中的文本"""
    # 优先处理Literal节点
    literal_value = get_literal_value(node)
    if literal_value is not None:
        return literal_value
    
    if not hasattr(node, 'position') or not node.position:
        return str(node)
    
    start_pos = node.position
    end_pos = node.position_end if hasattr(node, 'position_end') else None
    
    if not end_pos:
        return str(node)
    
    lines = code.split('\n')
    start_line = start_pos.line - 1
    end_line = end_pos.line - 1
    
    if start_line >= len(lines) or end_line >= len(lines):
        return str(node)
    
    if start_line == end_line:
        return lines[start_line][start_pos.column-1:end_pos.column]
    else:
        result = []
        for i in range(start_line, min(end_line+1, start_line+3)):
            if i == start_line:
                result.append(lines[i][start_pos.column-1:])
            elif i == end_line:
                result.append(lines[i][:end_pos.column])
            else:
                result.append(lines[i])
        if end_line > start_line + 2:
            result.append('...')
        return '\n'.join(result)

def preprocess_java_code(code):
    """预处理Java代码：移除注解和泛型语法"""
    code = re.sub(r'@\w+\s+', '', code)
    code = re.sub(r'@\w+\(.*?\)', '', code)
    code = re.sub(r'<.*?>', '', code)
    code = re.sub(r'//.*', '', code)
    code = re.sub(r'/\*.*?\*/', '', code, flags=re.DOTALL)
    return code

def parse_java_file(file_path):
    """解析Java文件，修复Literal节点处理"""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            original_code = f.read()
        
        processed_code = preprocess_java_code(original_code)
        tree = javalang.parse.parse(processed_code)
        fields = []
        
        for path, node in tree.filter(javalang.tree.FieldDeclaration):
            for declarator in node.declarators:
                default_value = None
                if declarator.initializer:
                    default_value = get_node_text(declarator.initializer, original_code)
                
                field_info = {
                    '字段名': declarator.name,
                    '类型': node.type.name,
                    '默认值': default_value,
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
    parser = argparse.ArgumentParser(description='Java字段提取（修复Literal节点处理）')
    parser.add_argument('--input_folder', required=True, help='Java文件所在文件夹')
    parser.add_argument('--output_file', required=True, help='输出的Excel文件名')
    args = parser.parse_args()
    
    process_java_folder(args.input_folder, args.output_file)
