import docx
from docx.opc.exceptions import PackageNotFoundError
import os
import html
import re
from docx import Document

class WordReader:
    """Word文档模板读取器"""
    
    def read_template(self, file_path):
        """读取Word模板并返回内容和变量列表"""
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"找不到文件: {file_path}")
            
        try:
            doc = Document(file_path)
        except PackageNotFoundError:
            raise ValueError(f"无法打开文件，可能不是有效的Word文档: {file_path}")
        
        content = []
        variables = set()  # 使用集合存储找到的所有变量
        
        # 遍历所有段落
        for para in doc.paragraphs:
            content.append(para.text)
            # 查找所有 {变量名} 格式的变量
            vars = re.findall(r'\{([^}]+)\}', para.text)
            variables.update(vars)
        
        return '\n'.join(content), list(variables)

    def read_template_html(self, file_path):
        """
        读取Word文档内容，并转为HTML格式
        支持{name}和{email}作为替换变量
        """
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"找不到文件: {file_path}")
            
        try:
            doc = docx.Document(file_path)
        except PackageNotFoundError:
            raise ValueError(f"无法打开文件，可能不是有效的Word文档: {file_path}")
        
        # 转换为HTML
        html_content = []
        
        for para in doc.paragraphs:
            if para.text.strip():
                # 处理段落样式
                style = ""
                if para.style.name.startswith('Heading'):
                    level = para.style.name[-1]
                    html_content.append(f"<h{level}>{html.escape(para.text)}</h{level}>")
                else:
                    # 处理段落中的格式
                    formatted_text = []
                    for run in para.runs:
                        text = html.escape(run.text)
                        if run.bold:
                            text = f"<strong>{text}</strong>"
                        if run.italic:
                            text = f"<em>{text}</em>"
                        if run.underline:
                            text = f"<u>{text}</u>"
                        formatted_text.append(text)
                    
                    html_content.append(f"<p>{''.join(formatted_text)}</p>")
        
        return "\n".join(html_content) 