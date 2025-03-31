import pandas as pd
import os

class ExcelReader:
    """Excel文件读取器"""
    
    def read_data(self, file_path):
        """
        读取Excel文件数据
        返回数据列表和列名列表
        """
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"找不到文件: {file_path}")
            
        try:
            # 读取Excel
            df = pd.read_excel(file_path)
            
            # 验证数据帧不为空
            if df.empty:
                raise ValueError("Excel文件中没有数据")
                
            # 将数据转换为列表字典
            data = df.to_dict(orient='records')
            
            # 获取列名
            columns = df.columns.tolist()
            
            return data, columns
            
        except Exception as e:
            raise ValueError(f"读取Excel文件时出错: {str(e)}") 