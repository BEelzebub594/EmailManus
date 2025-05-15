import os
import json
import shutil

class TemplateManager:
    def __init__(self):
        self.template_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "templates")
        # 确保模板目录存在
        if not os.path.exists(self.template_dir):
            os.makedirs(self.template_dir)
    
    def get_templates(self):
        """获取所有模板名称"""
        templates = []
        try:
            for file in os.listdir(self.template_dir):
                if file.endswith(".json"):
                    templates.append(os.path.splitext(file)[0])
        except Exception as e:
            print(f"获取模板列表出错: {str(e)}")
        return templates
    
    def get_template_content(self, template_name):
        """获取指定模板的内容"""
        try:
            template_path = os.path.join(self.template_dir, f"{template_name}.json")
            if os.path.exists(template_path):
                with open(template_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    return data.get("subject", ""), data.get("content", "")
        except Exception as e:
            print(f"读取模板内容出错: {str(e)}")
        return "", ""
    
    def save_template(self, name, subject, content):
        """保存模板"""
        try:
            template_path = os.path.join(self.template_dir, f"{name}.json")
            with open(template_path, "w", encoding="utf-8") as f:
                json.dump({"subject": subject, "content": content}, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            print(f"保存模板出错: {str(e)}")
            return False
    
    def delete_template(self, name):
        """删除模板"""
        try:
            template_path = os.path.join(self.template_dir, f"{name}.json")
            if os.path.exists(template_path):
                os.remove(template_path)
            return True
        except Exception as e:
            print(f"删除模板出错: {str(e)}")
            return False 