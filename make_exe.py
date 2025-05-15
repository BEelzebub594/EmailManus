import os
import shutil
import subprocess
import sys
from PIL import Image  # 添加PIL库用于图像处理

def jpg_to_ico(jpg_path, ico_path, size=(256, 256)):
    """将JPG图像转换为ICO图标"""
    try:
        if not os.path.exists(jpg_path):
            print(f"错误: 找不到图像文件 {jpg_path}")
            return False
            
        img = Image.open(jpg_path)
        img = img.resize(size, Image.LANCZOS)  # 重新调整大小
        img.save(ico_path, format='ICO')
        print(f"已成功将 {jpg_path} 转换为图标: {ico_path}")
        return True
    except Exception as e:
        print(f"转换图标时出错: {e}")
        return False

def clean_files():
    """清理不需要的文件和目录"""
    print("清理临时文件...")
    
    # 要删除的文件列表
    files_to_delete = [
        'create_sample_excel.py',
        'create_example_attachments.py',
        'create_package.py'
    ]
    
    # 要删除的目录列表
    dirs_to_delete = ['build', '__pycache__']
    
    # 删除文件
    for file in files_to_delete:
        if os.path.exists(file):
            os.remove(file)
            print(f"已删除: {file}")
    
    # 删除目录
    for dir_name in dirs_to_delete:
        if os.path.exists(dir_name):
            shutil.rmtree(dir_name)
            print(f"已删除目录: {dir_name}")
    
    # 清理core和ui中的__pycache__
    for root_dir in ['core', 'ui']:
        if os.path.exists(root_dir):
            pycache_path = os.path.join(root_dir, '__pycache__')
            if os.path.exists(pycache_path):
                shutil.rmtree(pycache_path)
                print(f"已删除目录: {pycache_path}")

def create_exe():
    """创建独立的EXE文件"""
    print("开始打包为EXE文件...")
    
    # 确保安装了pyinstaller
    try:
        subprocess.run([sys.executable, "-m", "pip", "install", "pyinstaller"], check=True)
        subprocess.run([sys.executable, "-m", "pip", "install", "pillow"], check=True)  # 确保安装PIL/Pillow
    except:
        print("PyInstaller或Pillow安装失败，请手动安装: pip install pyinstaller pillow")
        return False
    
    # 创建图标
    icon_path = "app_icon.ico"
    avatar_path = "头像.jpg"
    
    if os.path.exists(avatar_path):
        print(f"发现自定义头像: {avatar_path}")
        if not jpg_to_ico(avatar_path, icon_path):
            # 如果转换失败，尝试使用默认图标
            icon_path = "email_icon.ico" if os.path.exists("email_icon.ico") else ""
            print(f"图标转换失败，将使用默认图标: {icon_path}")
    else:
        # 如果找不到头像，尝试使用默认图标
        icon_path = "email_icon.ico" if os.path.exists("email_icon.ico") else ""
        print(f"警告: 找不到头像文件 {avatar_path}，将使用默认图标: {icon_path}")
    
    # 检查打赏图片
    donation_files = ["donation_wechat.jpg", "donation_alipay.jpg"]
    donation_data = []
    for file in donation_files:
        if os.path.exists(file):
            print(f"发现打赏图片: {file}")
            donation_data.append((file, file))
    
    # 添加额外数据文件
    additional_data = [
        ('templates', 'templates'),
        ('sample', 'sample'),
        ('attachments', 'attachments'),
        ('README.md', '.'),  # 添加README到根目录
    ]
    
    # 添加打赏图片
    for src, dest in donation_data:
        additional_data.append((src, '.'))
    
    # 如果有自定义头像，也添加
    if os.path.exists(avatar_path):
        additional_data.append((avatar_path, '.'))
    
    # 将数据列表转为命令行参数
    data_args = []
    for src, dest in additional_data:
        if os.path.exists(src):
            data_args.append(f"--add-data={src};{dest}")
    
    # 打包命令
    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--name=邮件群发助手",
        "--onefile",  # 创建单个EXE文件
        "--noconsole",  # 不显示控制台窗口
    ]
    
    # 添加数据文件参数
    cmd.extend(data_args)
    
    # 添加其他参数
    cmd.extend([
        "--hidden-import=PyQt5",
        "--hidden-import=PyQt5.QtWidgets",
        "--hidden-import=PyQt5.QtCore",
        "--hidden-import=PyQt5.QtGui",
        "--hidden-import=pandas",
        "--hidden-import=win32com",
        "--hidden-import=win32com.client",
        "--hidden-import=pythoncom",
        "--hidden-import=openpyxl",
        "--hidden-import=winreg",  # 添加winreg模块，用于注册表访问
        "--hidden-import=urllib.parse",  # 添加urllib.parse模块，用于URL编码
        f"--icon={icon_path}" if icon_path else "",
        "main.py"
    ])
    
    # 移除空的参数
    cmd = [x for x in cmd if x]
    
    try:
        subprocess.run(cmd, check=True)
        print("打包完成!")
        
        # 将生成的EXE文件复制到根目录
        shutil.copy2(
            os.path.join("dist", "邮件群发助手.exe"), 
            "邮件群发助手.exe"
        )
        print("已生成: 邮件群发助手.exe")
        
        # 复制启动脚本
        for script in ["setup.bat", "startup.bat"]:
            if os.path.exists(script):
                shutil.copy2(script, os.path.join("dist", script))
                print(f"已复制启动脚本: {script}")
        
        # 复制requirements.txt
        if os.path.exists("requirements.txt"):
            shutil.copy2("requirements.txt", os.path.join("dist", "requirements.txt"))
            print("已复制依赖文件: requirements.txt")
        
        return True
    except Exception as e:
        print(f"打包过程出错: {e}")
        return False

def main():
    """主函数"""
    print("=== 邮件群发助手打包工具 ===")
    print("\n自定义说明:")
    print("1. 将头像文件命名为'头像.jpg'放在根目录，可自动设置为程序图标")
    print("2. 添加打赏图片，命名为'donation_wechat.jpg'或'donation_alipay.jpg'")
    print("3. 完成后打包文件将位于dist目录和根目录\n")
    print("4. 当前支持的邮件客户端: Outlook, Foxmail, Thunderbird, Windows Mail, 网易邮箱大师, QQ邮箱\n")
    
    # 创建EXE
    if create_exe():
        # 清理文件
        clean_files()
        print("\n打包完成!")
        print("邮件群发助手.exe 已生成在当前目录和dist目录")
        print("可直接双击startup.bat启动程序")
    else:
        print("\n打包过程出错!")

if __name__ == "__main__":
    main() 