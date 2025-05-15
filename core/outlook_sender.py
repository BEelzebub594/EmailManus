import win32com.client
import pythoncom
import re
import time
import os
import sys
import subprocess
import traceback
import webbrowser
import glob
from pathlib import Path
import winreg

class EmailSender:
    """通用邮件发送器，支持多种邮件客户端"""
    
    # 支持的邮件客户端类型
    CLIENT_OUTLOOK = "outlook"
    CLIENT_FOXMAIL = "foxmail"
    CLIENT_THUNDERBIRD = "thunderbird"
    CLIENT_WINDOWS_MAIL = "windows_mail"
    CLIENT_NETEASE = "netease"  # 网易邮箱大师
    CLIENT_QQ_MAIL = "qq_mail"  # QQ邮箱客户端
    CLIENT_DEFAULT = "default"  # 系统默认邮件客户端
    
    def __init__(self, client_type=None):
        """初始化邮件发送器
        
        Args:
            client_type: 邮件客户端类型，默认为None(自动检测)
        """
        self.outlook = None
        self.mapi = None
        self.client_type = client_type  # 存储当前选择的客户端类型
        self.client_paths = {}  # 存储找到的客户端路径
        
        # 如果未指定客户端类型，尝试自动检测
        if not self.client_type:
            self.detect_available_clients()
            
            # 如果找到Outlook，默认使用Outlook
            if self.CLIENT_OUTLOOK in self.client_paths:
                self.client_type = self.CLIENT_OUTLOOK
            # 否则使用系统默认
            else:
                self.client_type = self.CLIENT_DEFAULT
    
    def detect_available_clients(self):
        """检测系统中可用的邮件客户端"""
        self.client_paths = {}
        
        # 检测Outlook
        try:
            # 使用WMI检测Outlook
            import win32com.client
            wmi = win32com.client.GetObject("winmgmts:")
            outlook_paths = wmi.ExecQuery("Select * from Win32_Process Where Name = 'OUTLOOK.EXE'")
            
            if len(outlook_paths) > 0:
                # Outlook正在运行
                self.client_paths[self.CLIENT_OUTLOOK] = "OUTLOOK.EXE"
            else:
                # 检查Outlook是否安装
                try:
                    outlook_key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, 
                                         r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE")
                    outlook_path = winreg.QueryValue(outlook_key, None)
                    if outlook_path:
                        self.client_paths[self.CLIENT_OUTLOOK] = outlook_path
                except:
                    try:
                        # 另一种检测方法
                        pythoncom.CoInitialize()
                        win32com.client.Dispatch("Outlook.Application")
                        self.client_paths[self.CLIENT_OUTLOOK] = "OUTLOOK.EXE"
                    except:
                        pass
        except:
            pass
            
        # 检测Foxmail
        try:
            # 常见的Foxmail安装路径
            foxmail_paths = [
                os.path.expandvars(r"%ProgramFiles%\Foxmail\Foxmail.exe"),
                os.path.expandvars(r"%ProgramFiles(x86)%\Foxmail\Foxmail.exe"),
                os.path.expandvars(r"%APPDATA%\Foxmail\Foxmail.exe"),
                r"C:\Program Files\Foxmail\Foxmail.exe",
                r"C:\Program Files (x86)\Foxmail\Foxmail.exe",
            ]
            
            # 尝试从注册表获取
            try:
                foxmail_key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, 
                                     r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Foxmail.exe")
                foxmail_path = winreg.QueryValue(foxmail_key, None)
                if foxmail_path:
                    foxmail_paths.insert(0, foxmail_path)  # 添加到列表开头
            except:
                pass
                
            # 检查所有可能的路径
            for path in foxmail_paths:
                if os.path.exists(path):
                    self.client_paths[self.CLIENT_FOXMAIL] = path
                    break
        except:
            pass
            
        # 检测Thunderbird
        try:
            # 常见的Thunderbird安装路径
            thunderbird_paths = [
                os.path.expandvars(r"%ProgramFiles%\Mozilla Thunderbird\thunderbird.exe"),
                os.path.expandvars(r"%ProgramFiles(x86)%\Mozilla Thunderbird\thunderbird.exe"),
                r"C:\Program Files\Mozilla Thunderbird\thunderbird.exe",
                r"C:\Program Files (x86)\Mozilla Thunderbird\thunderbird.exe",
            ]
            
            # 尝试从注册表获取
            try:
                thunderbird_key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, 
                                     r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\thunderbird.exe")
                thunderbird_path = winreg.QueryValue(thunderbird_key, None)
                if thunderbird_path:
                    thunderbird_paths.insert(0, thunderbird_path)  # 添加到列表开头
            except:
                pass
                
            # 检查所有可能的路径
            for path in thunderbird_paths:
                if os.path.exists(path):
                    self.client_paths[self.CLIENT_THUNDERBIRD] = path
                    break
        except:
            pass
            
        # 检测Windows Mail
        try:
            # Windows Mail是UWP应用，通过注册表检查是否可用
            mail_app_key_path = r"SOFTWARE\Classes\mailto\shell\open\command"
            try:
                mail_app_key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, mail_app_key_path)
                mail_command = winreg.QueryValue(mail_app_key, None)
                # 如果包含"microsoft.windowscommunicationsapps"则表示是Windows Mail
                if "microsoft.windowscommunicationsapps" in mail_command.lower():
                    self.client_paths[self.CLIENT_WINDOWS_MAIL] = "windows_mail"
            except:
                pass
        except:
            pass
            
        # 检测网易邮箱大师
        try:
            # 常见的网易邮箱大师安装路径
            netease_paths = [
                os.path.expandvars(r"%ProgramFiles%\NetEase\MailMaster\mailmaster.exe"),
                os.path.expandvars(r"%ProgramFiles(x86)%\NetEase\MailMaster\mailmaster.exe"),
                r"C:\Program Files\NetEase\MailMaster\mailmaster.exe",
                r"C:\Program Files (x86)\NetEase\MailMaster\mailmaster.exe",
            ]
            
            # 尝试从注册表获取
            try:
                netease_key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, 
                                     r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\mailmaster.exe")
                netease_path = winreg.QueryValue(netease_key, None)
                if netease_path:
                    netease_paths.insert(0, netease_path)  # 添加到列表开头
            except:
                pass
                
            # 检查所有可能的路径
            for path in netease_paths:
                if os.path.exists(path):
                    self.client_paths[self.CLIENT_NETEASE] = path
                    break
        except:
            pass
            
        # 检测QQ邮箱客户端
        try:
            # 常见的QQ邮箱客户端安装路径
            qq_mail_paths = [
                os.path.expandvars(r"%ProgramFiles%\Tencent\QQMail\QQMail.exe"),
                os.path.expandvars(r"%ProgramFiles(x86)%\Tencent\QQMail\QQMail.exe"),
                r"C:\Program Files\Tencent\QQMail\QQMail.exe",
                r"C:\Program Files (x86)\Tencent\QQMail\QQMail.exe",
            ]
            
            # 尝试从注册表获取
            try:
                qq_mail_key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, 
                                     r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\QQMail.exe")
                qq_mail_path = winreg.QueryValue(qq_mail_key, None)
                if qq_mail_path:
                    qq_mail_paths.insert(0, qq_mail_path)  # 添加到列表开头
            except:
                pass
                
            # 检查所有可能的路径
            for path in qq_mail_paths:
                if os.path.exists(path):
                    self.client_paths[self.CLIENT_QQ_MAIL] = path
                    break
        except:
            pass
            
        # 系统默认邮件客户端始终可用
        self.client_paths[self.CLIENT_DEFAULT] = "default"
        
        return self.client_paths
    
    def set_client(self, client_type):
        """设置要使用的邮件客户端"""
        if client_type in self.client_paths:
            self.client_type = client_type
            
            # 如果更改为Outlook，重置Outlook连接
            if client_type == self.CLIENT_OUTLOOK:
                self.outlook = None
                self.mapi = None
            
            return True
        else:
            return False
    
    def get_available_clients(self):
        """获取可用的邮件客户端列表"""
        if not self.client_paths:
            self.detect_available_clients()
            
        # 返回包含客户端友好名称的字典
        clients = {}
        for client, path in self.client_paths.items():
            if client == self.CLIENT_OUTLOOK:
                clients[client] = "Microsoft Outlook"
            elif client == self.CLIENT_FOXMAIL:
                clients[client] = "Foxmail"
            elif client == self.CLIENT_THUNDERBIRD:
                clients[client] = "Mozilla Thunderbird"
            elif client == self.CLIENT_WINDOWS_MAIL:
                clients[client] = "Windows Mail"
            elif client == self.CLIENT_NETEASE:
                clients[client] = "网易邮箱大师"
            elif client == self.CLIENT_QQ_MAIL:
                clients[client] = "QQ邮箱客户端"
            elif client == self.CLIENT_DEFAULT:
                clients[client] = "系统默认邮件客户端"
        
        return clients
    
    def connect_outlook(self):
        """连接到Outlook应用"""
        try:
            # 初始化COM环境
            pythoncom.CoInitialize()
            
            # 尝试方法1: 直接连接
            try:
                self.outlook = win32com.client.Dispatch("Outlook.Application")
                # 检查连接是否成功
                test = self.outlook.GetNamespace("MAPI")
                self.mapi = test
                return True
            except Exception as e:
                print(f"直接连接Outlook失败: {str(e)}")
                print(traceback.format_exc())
            
            # 尝试备用方案
            return False
        except Exception as e:
            print(f"连接Outlook出错: {str(e)}")
            print(traceback.format_exc())
            return False
    
    def get_sender_profiles(self):
        """获取可用的发件人邮箱列表"""
        profiles = []
        try:
            if not self.outlook:
                if not self.connect_outlook():
                    return profiles
            
            # 尝试获取账户信息
            try:
                if self.mapi:
                    accounts = self.mapi.Accounts
                    for i in range(1, accounts.Count + 1):
                        try:
                            account = accounts.Item(i)
                            profiles.append({
                                'name': account.DisplayName,
                                'email': account.SmtpAddress,
                                'account': account
                            })
                        except:
                            pass
            except Exception as e:
                print(f"获取账户列表失败: {str(e)}")
        except Exception as e:
            print(f"获取发件人配置文件出错: {str(e)}")
        
        return profiles
    
    def replace_variables(self, text, data):
        """替换文本中的变量"""
        if not text:
            return ""
        result = text
        for key, value in data.items():
            placeholder = f"{{{key}}}"
            if placeholder in result:
                # 确保value是字符串类型
                str_value = str(value) if value is not None else ""
                result = result.replace(placeholder, str_value)
        return result
    
    def find_attachments(self, attachment_pattern, data, custom_dir=None):
        """查找匹配的附件文件
        
        attachment_pattern: 附件名模式，可以包含变量，例如"合同_{姓名}.pdf"
        data: 当前行的数据字典
        custom_dir: 自定义附件目录，如果提供则优先在此目录中查找
        
        返回匹配的文件路径列表
        """
        if not attachment_pattern:
            return []
            
        # 替换附件模式中的变量
        pattern = self.replace_variables(attachment_pattern, data)
        
        # 添加通配符，使模式更灵活
        if "*" not in pattern and "?" not in pattern:
            pattern = f"*{pattern}*"
            
        # 查找所有匹配的文件
        matched_files = []
        
        # 设置搜索目录
        search_dirs = []
        
        # 如果提供了自定义目录，只在自定义目录中搜索
        if custom_dir and os.path.exists(custom_dir):
            search_dirs.append(custom_dir)
        else:
            # 只有在没有提供自定义目录时，才使用默认目录
            default_dirs = [
                os.getcwd(),
                os.path.join(os.getcwd(), "attachments"),
                os.path.join(os.getcwd(), "sample")
            ]
            
            # 添加默认目录
            for directory in default_dirs:
                if os.path.exists(directory):
                    search_dirs.append(directory)
        
        # 在所有目录中查找匹配的文件
        for directory in search_dirs:
            # 在当前目录中直接搜索
            direct_match = os.path.join(directory, pattern)
            matched_files.extend(glob.glob(direct_match))
            
            # 也在子目录中搜索
            nested_match = os.path.join(directory, "**", pattern)
            matched_files.extend(glob.glob(nested_match, recursive=True))
        
        # 去重
        matched_files = list(set(matched_files))
                
        # 打印调试信息
        if matched_files:
            print(f"找到附件: {matched_files}")
        else:
            print(f"未找到匹配的附件。模式: {pattern}, 搜索目录: {search_dirs}")
            
        return matched_files
    
    def create_mail_directly(self, to_address, subject, body, auto_send=False, attachments=None):
        """不使用Outlook直接创建邮件"""
        try:
            # 检查当前选择的邮件客户端
            if self.client_type == self.CLIENT_FOXMAIL:
                return self.create_mail_foxmail(to_address, subject, body, attachments)
            elif self.client_type == self.CLIENT_THUNDERBIRD:
                return self.create_mail_thunderbird(to_address, subject, body, attachments)
            elif self.client_type == self.CLIENT_WINDOWS_MAIL:
                return self.create_mail_windows_mail(to_address, subject, body, attachments)
            elif self.client_type == self.CLIENT_NETEASE:
                return self.create_mail_netease(to_address, subject, body, attachments)
            elif self.client_type == self.CLIENT_QQ_MAIL:
                return self.create_mail_qq_mail(to_address, subject, body, attachments)
            elif self.client_type == self.CLIENT_DEFAULT:
                return self.create_mail_default(to_address, subject, body, attachments)
            else:
                # 默认使用HTML预览方式
                return self.create_mail_html_preview(to_address, subject, body, auto_send, attachments)
        except Exception as e:
            print(f"创建邮件失败: {str(e)}")
            print(traceback.format_exc())
            # 失败后尝试使用HTML预览
            return self.create_mail_html_preview(to_address, subject, body, auto_send, attachments)
    
    def create_mail_html_preview(self, to_address, subject, body, auto_send=False, attachments=None):
        """创建HTML邮件预览"""
        try:
            # 创建一个HTML文件
            import tempfile
            html_file = os.path.join(tempfile.gettempdir(), f"email_{int(time.time())}.html")
            
            # 为HTML格式修正换行符
            body_html = body
            if "<br>" not in body and "<p>" not in body:
                # 如果邮件内容不包含HTML标签，将换行符转换为<br>
                body_html = body.replace("\n", "<br>\n")
            
            # 准备附件信息
            attachments_html = ""
            if attachments and len(attachments) > 0:
                attachments_html = "<div class='attachments'><p><strong>附件:</strong></p><ul>"
                for attachment in attachments:
                    file_name = os.path.basename(attachment)
                    attachments_html += f"<li>{file_name}</li>"
                attachments_html += "</ul></div>"
            
            with open(html_file, "w", encoding="utf-8") as f:
                html_content = f"""
                <!DOCTYPE html>
                <html>
                <head>
                    <meta charset="utf-8">
                    <title>{subject}</title>
                    <style>
                        body {{ font-family: Arial, sans-serif; padding: 20px; line-height: 1.6; }}
                        .email-container {{ border: 1px solid #ddd; padding: 20px; }}
                        .header {{ margin-bottom: 20px; border-bottom: 1px solid #eee; padding-bottom: 10px; }}
                        .content {{ white-space: pre-wrap; }}
                        .footer {{ margin-top: 20px; border-top: 1px solid #eee; padding-top: 10px; color: #777; }}
                        .warning {{ color: red; font-weight: bold; margin-top: 20px; padding: 10px; border: 1px solid red; }}
                        .attachments {{ margin-top: 15px; background-color: #f5f5f5; padding: 10px; border-radius: 5px; }}
                        .attachments ul {{ margin: 5px 0; padding-left: 20px; }}
                    </style>
                </head>
                <body>
                    <div class="email-container">
                        <div class="header">
                            <p><strong>To:</strong> {to_address}</p>
                            <p><strong>Subject:</strong> {subject}</p>
                        </div>
                        <div class="content">
                            {body_html}
                        </div>
                        {attachments_html}
                        <div class="footer">
                            <p>此邮件通过邮件群发助手生成</p>
                        </div>
                        {"<div class='warning'>注意：自动发送模式已开启，但无法连接到邮件客户端。这是邮件预览，需要手动发送。</div>" if auto_send else ""}
                    </div>
                </body>
                </html>
                """
                f.write(html_content)
            
            # 打开HTML文件
            if os.path.exists(html_file):
                webbrowser.open('file://' + os.path.abspath(html_file))
                return True
            else:
                print(f"错误：HTML文件未创建: {html_file}")
                return False
        except Exception as e:
            print(f"创建邮件HTML预览失败: {str(e)}")
            print(traceback.format_exc())
            return False
    
    def create_mail_foxmail(self, to_address, subject, body, attachments=None):
        """使用Foxmail创建邮件"""
        try:
            if self.CLIENT_FOXMAIL not in self.client_paths:
                print("错误：未找到Foxmail路径")
                return self.create_mail_html_preview(to_address, subject, body, False, attachments)
            
            foxmail_path = self.client_paths[self.CLIENT_FOXMAIL]
            
            # 准备mailto链接参数
            mailto_params = {
                "to": to_address,
                "subject": subject,
                "body": body
            }
            
            # 构建mailto URL
            import urllib.parse
            mailto_url = f"mailto:{mailto_params['to']}?subject={urllib.parse.quote(mailto_params['subject'])}&body={urllib.parse.quote(mailto_params['body'])}"
            
            # 启动Foxmail并传递mailto链接
            try:
                # 尝试使用启动参数方式
                cmd = f'"{foxmail_path}" "{mailto_url}"'
                subprocess.Popen(cmd, shell=True)
                
                # 如果有附件，提示用户手动添加
                if attachments and len(attachments) > 0:
                    import tempfile
                    attachments_file = os.path.join(tempfile.gettempdir(), f"attachments_{int(time.time())}.txt")
                    with open(attachments_file, "w", encoding="utf-8") as f:
                        f.write("邮件群发助手附件列表：\n\n")
                        for attachment in attachments:
                            f.write(f"{attachment}\n")
                    
                    # 打开附件列表文件
                    os.startfile(attachments_file)
                    
                    # 显示提示窗口（如果需要）
                    # 使用subprocess打开一个简单的MessageBox
                    msg_cmd = f'powershell -Command "[System.Windows.Forms.MessageBox]::Show(\'请手动添加以下附件:\\n{len(attachments)}个附件\\n详细列表已打开\', \'邮件群发助手\', \'OK\', \'Information\')"'
                    subprocess.Popen(msg_cmd, shell=True)
                
                return True
            except Exception as e:
                print(f"启动Foxmail失败: {str(e)}")
                # 尝试备用方案
                webbrowser.open(mailto_url)
                return True
                
        except Exception as e:
            print(f"使用Foxmail创建邮件失败: {str(e)}")
            print(traceback.format_exc())
            # 失败时使用HTML预览
            return self.create_mail_html_preview(to_address, subject, body, False, attachments)
    
    def create_mail_thunderbird(self, to_address, subject, body, attachments=None):
        """使用Thunderbird创建邮件"""
        try:
            if self.CLIENT_THUNDERBIRD not in self.client_paths:
                print("错误：未找到Thunderbird路径")
                return self.create_mail_html_preview(to_address, subject, body, False, attachments)
            
            thunderbird_path = self.client_paths[self.CLIENT_THUNDERBIRD]
            
            # 准备mailto链接参数
            mailto_params = {
                "to": to_address,
                "subject": subject,
                "body": body
            }
            
            # 构建mailto URL
            import urllib.parse
            mailto_url = f"mailto:{mailto_params['to']}?subject={urllib.parse.quote(mailto_params['subject'])}&body={urllib.parse.quote(mailto_params['body'])}"
            
            # 启动Thunderbird并传递mailto链接
            try:
                # 尝试使用启动参数方式
                cmd = f'"{thunderbird_path}" -compose "{mailto_url}"'
                subprocess.Popen(cmd, shell=True)
                
                # 如果有附件，提示用户手动添加
                if attachments and len(attachments) > 0:
                    import tempfile
                    attachments_file = os.path.join(tempfile.gettempdir(), f"attachments_{int(time.time())}.txt")
                    with open(attachments_file, "w", encoding="utf-8") as f:
                        f.write("邮件群发助手附件列表：\n\n")
                        for attachment in attachments:
                            f.write(f"{attachment}\n")
                    
                    # 打开附件列表文件
                    os.startfile(attachments_file)
                    
                    # 显示提示窗口
                    msg_cmd = f'powershell -Command "[System.Windows.Forms.MessageBox]::Show(\'请手动添加以下附件:\\n{len(attachments)}个附件\\n详细列表已打开\', \'邮件群发助手\', \'OK\', \'Information\')"'
                    subprocess.Popen(msg_cmd, shell=True)
                
                return True
            except Exception as e:
                print(f"启动Thunderbird失败: {str(e)}")
                # 尝试备用方案
                webbrowser.open(mailto_url)
                return True
                
        except Exception as e:
            print(f"使用Thunderbird创建邮件失败: {str(e)}")
            print(traceback.format_exc())
            # 失败时使用HTML预览
            return self.create_mail_html_preview(to_address, subject, body, False, attachments)
    
    def create_mail_windows_mail(self, to_address, subject, body, attachments=None):
        """使用Windows Mail创建邮件"""
        try:
            # Windows Mail是UWP应用，可以通过URI scheme启动
            # 准备mailto链接参数
            mailto_params = {
                "to": to_address,
                "subject": subject,
                "body": body
            }
            
            # 构建mailto URL
            import urllib.parse
            mailto_url = f"mailto:{mailto_params['to']}?subject={urllib.parse.quote(mailto_params['subject'])}&body={urllib.parse.quote(mailto_params['body'])}"
            
            # 尝试使用Windows Mail应用打开
            try:
                # 使用启动默认邮件应用的方式，Windows会使用默认的Mail应用
                os.startfile(mailto_url)
                
                # 如果有附件，提示用户手动添加
                if attachments and len(attachments) > 0:
                    import tempfile
                    attachments_file = os.path.join(tempfile.gettempdir(), f"attachments_{int(time.time())}.txt")
                    with open(attachments_file, "w", encoding="utf-8") as f:
                        f.write("邮件群发助手附件列表：\n\n")
                        for attachment in attachments:
                            f.write(f"{attachment}\n")
                    
                    # 打开附件列表文件
                    os.startfile(attachments_file)
                    
                    # 显示提示窗口
                    msg_cmd = f'powershell -Command "[System.Windows.Forms.MessageBox]::Show(\'请手动添加以下附件:\\n{len(attachments)}个附件\\n详细列表已打开\', \'邮件群发助手\', \'OK\', \'Information\')"'
                    subprocess.Popen(msg_cmd, shell=True)
                
                return True
            except Exception as e:
                print(f"启动Windows Mail失败: {str(e)}")
                # 尝试备用方案
                webbrowser.open(mailto_url)
                return True
                
        except Exception as e:
            print(f"使用Windows Mail创建邮件失败: {str(e)}")
            print(traceback.format_exc())
            # 失败时使用HTML预览
            return self.create_mail_html_preview(to_address, subject, body, False, attachments)
    
    def create_mail_netease(self, to_address, subject, body, attachments=None):
        """使用网易邮箱大师创建邮件"""
        try:
            if self.CLIENT_NETEASE not in self.client_paths:
                print("错误：未找到网易邮箱大师路径")
                return self.create_mail_html_preview(to_address, subject, body, False, attachments)
            
            netease_path = self.client_paths[self.CLIENT_NETEASE]
            
            # 准备mailto链接参数
            mailto_params = {
                "to": to_address,
                "subject": subject,
                "body": body
            }
            
            # 构建mailto URL
            import urllib.parse
            mailto_url = f"mailto:{mailto_params['to']}?subject={urllib.parse.quote(mailto_params['subject'])}&body={urllib.parse.quote(mailto_params['body'])}"
            
            # 启动网易邮箱大师并传递mailto链接
            try:
                # 尝试使用启动参数方式
                cmd = f'"{netease_path}" "{mailto_url}"'
                subprocess.Popen(cmd, shell=True)
                
                # 如果有附件，提示用户手动添加
                if attachments and len(attachments) > 0:
                    import tempfile
                    attachments_file = os.path.join(tempfile.gettempdir(), f"attachments_{int(time.time())}.txt")
                    with open(attachments_file, "w", encoding="utf-8") as f:
                        f.write("邮件群发助手附件列表：\n\n")
                        for attachment in attachments:
                            f.write(f"{attachment}\n")
                    
                    # 打开附件列表文件
                    os.startfile(attachments_file)
                    
                    # 显示提示窗口
                    msg_cmd = f'powershell -Command "[System.Windows.Forms.MessageBox]::Show(\'请手动添加以下附件:\\n{len(attachments)}个附件\\n详细列表已打开\', \'邮件群发助手\', \'OK\', \'Information\')"'
                    subprocess.Popen(msg_cmd, shell=True)
                
                return True
            except Exception as e:
                print(f"启动网易邮箱大师失败: {str(e)}")
                # 尝试备用方案
                webbrowser.open(mailto_url)
                return True
                
        except Exception as e:
            print(f"使用网易邮箱大师创建邮件失败: {str(e)}")
            print(traceback.format_exc())
            # 失败时使用HTML预览
            return self.create_mail_html_preview(to_address, subject, body, False, attachments)
    
    def create_mail_qq_mail(self, to_address, subject, body, attachments=None):
        """使用QQ邮箱客户端创建邮件"""
        try:
            if self.CLIENT_QQ_MAIL not in self.client_paths:
                print("错误：未找到QQ邮箱客户端路径")
                return self.create_mail_html_preview(to_address, subject, body, False, attachments)
            
            qq_mail_path = self.client_paths[self.CLIENT_QQ_MAIL]
            
            # 准备mailto链接参数
            mailto_params = {
                "to": to_address,
                "subject": subject,
                "body": body
            }
            
            # 构建mailto URL
            import urllib.parse
            mailto_url = f"mailto:{mailto_params['to']}?subject={urllib.parse.quote(mailto_params['subject'])}&body={urllib.parse.quote(mailto_params['body'])}"
            
            # 启动QQ邮箱客户端并传递mailto链接
            try:
                # 尝试使用启动参数方式
                cmd = f'"{qq_mail_path}" "{mailto_url}"'
                subprocess.Popen(cmd, shell=True)
                
                # 如果有附件，提示用户手动添加
                if attachments and len(attachments) > 0:
                    import tempfile
                    attachments_file = os.path.join(tempfile.gettempdir(), f"attachments_{int(time.time())}.txt")
                    with open(attachments_file, "w", encoding="utf-8") as f:
                        f.write("邮件群发助手附件列表：\n\n")
                        for attachment in attachments:
                            f.write(f"{attachment}\n")
                    
                    # 打开附件列表文件
                    os.startfile(attachments_file)
                    
                    # 显示提示窗口
                    msg_cmd = f'powershell -Command "[System.Windows.Forms.MessageBox]::Show(\'请手动添加以下附件:\\n{len(attachments)}个附件\\n详细列表已打开\', \'邮件群发助手\', \'OK\', \'Information\')"'
                    subprocess.Popen(msg_cmd, shell=True)
                
                return True
            except Exception as e:
                print(f"启动QQ邮箱客户端失败: {str(e)}")
                # 尝试备用方案
                webbrowser.open(mailto_url)
                return True
                
        except Exception as e:
            print(f"使用QQ邮箱客户端创建邮件失败: {str(e)}")
            print(traceback.format_exc())
            # 失败时使用HTML预览
            return self.create_mail_html_preview(to_address, subject, body, False, attachments)
    
    def create_mail_default(self, to_address, subject, body, attachments=None):
        """使用系统默认邮件客户端创建邮件"""
        try:
            # 准备mailto链接参数
            mailto_params = {
                "to": to_address,
                "subject": subject,
                "body": body
            }
            
            # 构建mailto URL
            import urllib.parse
            mailto_url = f"mailto:{mailto_params['to']}?subject={urllib.parse.quote(mailto_params['subject'])}&body={urllib.parse.quote(mailto_params['body'])}"
            
            # 使用webbrowser打开mailto链接
            webbrowser.open(mailto_url)
            
            # 如果有附件，提示用户手动添加
            if attachments and len(attachments) > 0:
                import tempfile
                attachments_file = os.path.join(tempfile.gettempdir(), f"attachments_{int(time.time())}.txt")
                with open(attachments_file, "w", encoding="utf-8") as f:
                    f.write("邮件群发助手附件列表：\n\n")
                    for attachment in attachments:
                        f.write(f"{attachment}\n")
                
                # 打开附件列表文件
                os.startfile(attachments_file)
                
                # 显示提示窗口
                msg_cmd = f'powershell -Command "[System.Windows.Forms.MessageBox]::Show(\'请手动添加以下附件:\\n{len(attachments)}个附件\\n详细列表已打开\', \'邮件群发助手\', \'OK\', \'Information\')"'
                subprocess.Popen(msg_cmd, shell=True)
            
            return True
        except Exception as e:
            print(f"使用默认邮件客户端创建邮件失败: {str(e)}")
            print(traceback.format_exc())
            # 失败时使用HTML预览
            return self.create_mail_html_preview(to_address, subject, body, False, attachments)
            
    def send_batch_emails(self, data_list, to_column, subject_template, body_template, sender_email=None, auto_send=False, attachment_pattern=None, attachment_dir=None):
        """批量发送邮件"""
        # 只有当选择Outlook时才连接Outlook
        outlook_connected = False
        if self.client_type == self.CLIENT_OUTLOOK:
            outlook_connected = self.connect_outlook()
            if not outlook_connected:
                print("警告: 无法连接到Outlook，将尝试使用替代方法创建邮件。")
                if auto_send:
                    print("警告: 自动发送功能需要连接到Outlook，将改为预览模式。")
                    auto_send = False
        else:
            # 非Outlook客户端不支持自动发送
            if auto_send:
                print(f"警告: {self.client_type}客户端不支持自动发送，将改为预览模式。")
                auto_send = False
        
        sent_count = 0
        for data in data_list:
            try:
                # 获取收件人
                to_address = data.get(to_column, "")
                if not to_address:
                    continue
                
                # 替换变量
                subject = self.replace_variables(subject_template, data)
                body = self.replace_variables(body_template, data)
                
                # 查找附件
                attachments = self.find_attachments(attachment_pattern, data, attachment_dir) if attachment_pattern else []
                
                # 过滤有效附件（文件必须存在且大小大于0）
                valid_attachments = []
                for attachment_path in attachments:
                    try:
                        if os.path.exists(attachment_path):
                            file_size = os.path.getsize(attachment_path)
                            if file_size > 0:
                                valid_attachments.append(attachment_path)
                                print(f"有效附件: {attachment_path}, 大小: {file_size} 字节")
                            else:
                                print(f"忽略空文件附件: {attachment_path}")
                        else:
                            print(f"附件文件不存在: {attachment_path}")
                    except Exception as e:
                        print(f"检查附件时出错: {attachment_path}, 错误: {str(e)}")
                
                if len(valid_attachments) != len(attachments):
                    print(f"注意: 共找到{len(attachments)}个附件，但只有{len(valid_attachments)}个有效")
                
                if outlook_connected and self.client_type == self.CLIENT_OUTLOOK:
                    try:
                        # 使用Outlook创建邮件
                        mail = self.outlook.CreateItem(0)  # 0: olMailItem
                        mail.To = to_address
                        mail.Subject = subject
                        
                        # 确保正确的HTML格式
                        if not body.startswith("<html>"):
                            # 将换行符转换为HTML换行
                            if "<br>" not in body and "<p>" not in body:
                                html_body = body.replace("\n", "<br>\n")
                            else:
                                html_body = body
                                
                            # 添加HTML头和尾
                            mail.HTMLBody = f"""
                            <html>
                            <head>
                            <style>
                            body {{ font-family: Arial, sans-serif; line-height: 1.6; }}
                            </style>
                            </head>
                            <body>
                            {html_body}
                            </body>
                            </html>
                            """
                        else:
                            mail.HTMLBody = body
                        
                        # 添加附件
                        for attachment_path in valid_attachments:
                            try:
                                mail.Attachments.Add(attachment_path)
                                print(f"成功添加附件: {attachment_path}")
                            except Exception as e:
                                print(f"添加附件失败: {attachment_path}, 错误: {str(e)}")
                        
                        # 设置发件人
                        if sender_email:
                            try:
                                profiles = self.get_sender_profiles()
                                for profile in profiles:
                                    if profile['email'] == sender_email:
                                        mail.SendUsingAccount = profile['account']
                                        break
                            except Exception as e:
                                print(f"设置发件人账户失败: {str(e)}")
                        
                        # 根据选项决定显示还是直接发送
                        if auto_send:
                            mail.Send()
                            print(f"已自动发送邮件到: {to_address}, 附件数量: {len(valid_attachments)}")
                        else:
                            mail.Display()  # 显示邮件供用户确认
                            print(f"已创建邮件预览: {to_address}, 附件数量: {len(valid_attachments)}")
                        
                        sent_count += 1
                    except Exception as e:
                        print(f"Outlook创建邮件失败，尝试备用方法: {str(e)}")
                        print(traceback.format_exc())
                        success = self.create_mail_directly(to_address, subject, body, auto_send, valid_attachments)
                        if success:
                            sent_count += 1
                else:
                    # 使用替代方法创建邮件
                    success = self.create_mail_directly(to_address, subject, body, auto_send, valid_attachments)
                    if success:
                        sent_count += 1
                
                # 短暂暂停，避免响应问题
                time.sleep(0.5)
                
            except Exception as e:
                print(f"创建邮件出错: {str(e)}")
                print(traceback.format_exc())
        
        return sent_count 