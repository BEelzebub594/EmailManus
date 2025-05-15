import sys
import os

try:
    from PyQt5.QtWidgets import QApplication
    from PyQt5.QtGui import QIcon
except ImportError:
    print("错误: 缺少PyQt5模块，请确保它已正确安装")
    sys.exit(1)

from ui.app_ui import EmailManusApp

if __name__ == "__main__":
    # 确保我们可以正确找到资源文件
    if getattr(sys, 'frozen', False):
        # 如果是打包的可执行文件
        application_path = os.path.dirname(sys.executable)
    else:
        # 如果是脚本运行
        application_path = os.path.dirname(os.path.abspath(__file__))
        
    os.chdir(application_path)  # 切换到应用程序所在目录
    
    app = QApplication(sys.argv)
    
    # 设置应用程序图标
    app_icon_path = "app_icon.ico"
    avatar_path = "头像.jpg"
    
    if os.path.exists(app_icon_path):
        app.setWindowIcon(QIcon(app_icon_path))
    elif os.path.exists(avatar_path):
        app.setWindowIcon(QIcon(avatar_path))
    elif os.path.exists("email_icon.ico"):
        app.setWindowIcon(QIcon("email_icon.ico"))
    
    window = EmailManusApp()
    window.show()
    sys.exit(app.exec_()) 