import pandas as pd

class ExcelReader:
    def get_sheet_names(self, file_path):
        """获取Excel中的所有Sheet名称"""
        try:
            xl = pd.ExcelFile(file_path)
            return xl.sheet_names
        except Exception as e:
            print(f"读取Excel sheet列表出错: {str(e)}")
            return []
    
    def get_column_names(self, file_path, sheet_name):
        """获取指定Sheet中的列名"""
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            return df.columns.tolist()
        except Exception as e:
            print(f"读取Excel列名出错: {str(e)}")
            return []
    
    def read_data(self, file_path, sheet_name):
        """读取指定Sheet中的数据"""
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            # 转换为字典列表
            return df.to_dict('records')
        except Exception as e:
            print(f"读取Excel数据出错: {str(e)}")
            return [] 