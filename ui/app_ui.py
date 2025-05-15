from PyQt5.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                           QPushButton, QLabel, QComboBox, QFileDialog, 
                           QTextEdit, QTabWidget, QLineEdit, QMessageBox, QListWidget,
                           QInputDialog, QProgressBar, QApplication, QCheckBox, QFrame,
                           QSplitter, QGroupBox, QScrollArea)
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QFont, QIcon, QColor, QPalette, QPixmap
from core.excel_reader import ExcelReader
from core.template_manager import TemplateManager
from core.outlook_sender import EmailSender
import os
import sys
import traceback

def resource_path(relative_path):
    """获取资源的绝对路径，用于处理打包后的资源访问"""
    try:
        # PyInstaller创建临时文件夹并将路径存储在_MEIPASS中
        base_path = getattr(sys, '_MEIPASS', os.path.abspath("."))
        return os.path.join(base_path, relative_path)
    except Exception:
        return os.path.join(os.path.abspath("."), relative_path)

class EmailManusApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("邮件群发助手")
        self.setMinimumSize(900, 650)
        
        # 设置应用程序图标
        avatar_path = resource_path("头像.jpg")
        app_icon_path = resource_path("app_icon.ico")
        
        if os.path.exists(app_icon_path):
            self.setWindowIcon(QIcon(app_icon_path))
        elif os.path.exists(avatar_path):
            self.setWindowIcon(QIcon(avatar_path))
        else:
            # 尝试使用默认图标
            default_icon = resource_path("email_icon.ico")
            if os.path.exists(default_icon):
                self.setWindowIcon(QIcon(default_icon))
        
        # 设置应用程序样式
        self.setStyleSheet("""
            QMainWindow {
                background-color: #F5F6FA;
            }
            QTabWidget::pane {
                border: 1px solid #CCCCCC;
                background-color: white;
                border-radius: 5px;
            }
            QTabBar::tab {
                background-color: #E8E8E8;
                color: #444444;
                border: 1px solid #C4C4C4;
                border-bottom-color: #CCCCCC;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
                min-width: 6ex;
                max-width: 14ex;
                padding: 6px 10px;
                font-size: 10pt;
                font-weight: bold;
            }
            QTabBar::tab:selected {
                background-color: white;
                color: #3498DB;
                border-bottom-color: white;
            }
            QTabBar::tab:!selected {
                margin-top: 2px;
            }
            QPushButton {
                background-color: #3498DB;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 6px 12px;
                font-weight: bold;
                font-size: 9.5pt;
            }
            QPushButton:hover {
                background-color: #2980B9;
            }
            QPushButton:pressed {
                background-color: #1F618D;
            }
            QLineEdit, QComboBox, QTextEdit {
                border: 1px solid #CCC;
                border-radius: 4px;
                padding: 4px;
                background-color: white;
                font-size: 9.5pt;
            }
            QComboBox {
                min-height: 22px;
            }
            QGroupBox {
                font-weight: bold;
                border: 1px solid #CCC;
                border-radius: 5px;
                margin-top: 10px;
                padding-top: 16px;
                font-size: 10pt;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
            }
            QLabel {
                color: #333;
                font-size: 9.5pt;
            }
        """)
        
        self.excel_reader = ExcelReader()
        self.template_manager = TemplateManager()
        self.outlook_sender = EmailSender()
        
        self.init_ui()
        
        # 启动时自动尝试连接邮件客户端并获取账户列表
        QTimer.singleShot(500, self.load_sender_accounts)
    
    def init_ui(self):
        main_widget = QWidget()
        main_layout = QVBoxLayout()
        
        # 创建选项卡
        tabs = QTabWidget()
        
        # 发送邮件选项卡
        send_tab = QWidget()
        send_layout = QVBoxLayout()
        
        # Excel文件选择组
        excel_group = QGroupBox("数据源")
        excel_group_layout = QVBoxLayout()
        
        excel_layout = QHBoxLayout()
        excel_layout.addWidget(QLabel("Excel文件:"))
        self.excel_path = QLineEdit()
        excel_layout.addWidget(self.excel_path, 1)
        excel_btn = QPushButton("浏览...")
        excel_btn.clicked.connect(self.browse_excel)
        excel_layout.addWidget(excel_btn)
        excel_group_layout.addLayout(excel_layout)
        
        sheet_layout = QHBoxLayout()
        sheet_layout.addWidget(QLabel("Sheet:"))
        self.sheet_combo = QComboBox()
        sheet_layout.addWidget(self.sheet_combo, 1)
        load_btn = QPushButton("加载")
        load_btn.clicked.connect(self.load_excel_data)
        sheet_layout.addWidget(load_btn)
        excel_group_layout.addLayout(sheet_layout)
        
        email_layout = QHBoxLayout()
        email_layout.addWidget(QLabel("收件人邮箱列:"))
        self.email_column_combo = QComboBox()
        email_layout.addWidget(self.email_column_combo, 1)
        excel_group_layout.addLayout(email_layout)
        
        excel_group.setLayout(excel_group_layout)
        send_layout.addWidget(excel_group)
        
        # 邮件设置组
        mail_group = QGroupBox("邮件设置")
        mail_group_layout = QVBoxLayout()
        
        # 邮件客户端选择
        client_layout = QHBoxLayout()
        client_layout.addWidget(QLabel("邮件客户端:"))
        self.client_combo = QComboBox()
        client_layout.addWidget(self.client_combo, 1)
        refresh_client_btn = QPushButton("检测")
        refresh_client_btn.clicked.connect(self.detect_email_clients)
        client_layout.addWidget(refresh_client_btn)
        mail_group_layout.addLayout(client_layout)
        
        # 发件人选择
        sender_layout = QHBoxLayout()
        sender_layout.addWidget(QLabel("发件人邮箱:"))
        self.sender_combo = QComboBox()
        sender_layout.addWidget(self.sender_combo, 1)
        refresh_btn = QPushButton("刷新")
        refresh_btn.clicked.connect(self.load_sender_accounts)
        sender_layout.addWidget(refresh_btn)
        mail_group_layout.addLayout(sender_layout)
        
        # 模板选择
        template_layout = QHBoxLayout()
        template_layout.addWidget(QLabel("邮件模板:"))
        self.template_combo = QComboBox()
        self.template_combo.currentIndexChanged.connect(self.load_template)
        template_layout.addWidget(self.template_combo, 1)
        mail_group_layout.addLayout(template_layout)
        
        mail_group.setLayout(mail_group_layout)
        send_layout.addWidget(mail_group)
        
        # 预览和变量区域
        preview_group = QGroupBox("邮件内容")
        preview_layout = QVBoxLayout()
        
        splitter = QSplitter(Qt.Horizontal)
        
        # 变量列表
        var_widget = QWidget()
        var_layout = QVBoxLayout(var_widget)
        var_layout.addWidget(QLabel("可用变量:"))
        self.var_list = QListWidget()
        self.var_list.setStyleSheet("border: 1px solid #CCC; border-radius: 4px;")
        var_layout.addWidget(self.var_list)
        var_widget.setLayout(var_layout)
        splitter.addWidget(var_widget)
        
        # 邮件预览
        mail_widget = QWidget()
        mail_layout = QVBoxLayout(mail_widget)
        
        mail_subject_layout = QHBoxLayout()
        mail_subject_layout.addWidget(QLabel("主题:"))
        self.mail_subject = QLineEdit()
        mail_subject_layout.addWidget(self.mail_subject)
        mail_layout.addLayout(mail_subject_layout)
        
        self.mail_content = QTextEdit()
        mail_layout.addWidget(self.mail_content)
        
        # 添加附件选择
        attachment_layout = QHBoxLayout()
        attachment_layout.addWidget(QLabel("附件模式:"))
        self.attachment_pattern = QLineEdit()
        self.attachment_pattern.setPlaceholderText("例如: *{姓名}*.pdf 或 合同_{订单号}.docx")
        attachment_layout.addWidget(self.attachment_pattern)
        attachment_hint = QLabel("(可使用变量和通配符)")
        attachment_hint.setStyleSheet("color: #555;")
        attachment_layout.addWidget(attachment_hint)
        mail_layout.addLayout(attachment_layout)
        
        # 添加附件目录选择
        attachment_dir_layout = QHBoxLayout()
        attachment_dir_layout.addWidget(QLabel("附件目录:"))
        self.attachment_dir = QLineEdit()
        self.attachment_dir.setPlaceholderText("默认在当前目录、附件目录和sample目录中查找")
        attachment_dir_layout.addWidget(self.attachment_dir)
        attachment_dir_btn = QPushButton("浏览...")
        attachment_dir_btn.clicked.connect(self.browse_attachment_dir)
        attachment_dir_layout.addWidget(attachment_dir_btn)
        mail_layout.addLayout(attachment_dir_layout)
        
        mail_widget.setLayout(mail_layout)
        splitter.addWidget(mail_widget)
        
        # 设置分割比例
        splitter.setSizes([200, 600])
        preview_layout.addWidget(splitter)
        
        preview_group.setLayout(preview_layout)
        send_layout.addWidget(preview_group)
        
        # 发送按钮和控制区域
        control_group = QGroupBox("发送选项")
        control_layout = QVBoxLayout()
        
        # 自动发送选项
        auto_send_layout = QHBoxLayout()
        self.auto_send_checkbox = QCheckBox("自动发送邮件（不预览直接发送）")
        font = QFont("Arial", 10)
        font.setBold(True)
        self.auto_send_checkbox.setFont(font)
        auto_send_layout.addWidget(self.auto_send_checkbox)
        control_layout.addLayout(auto_send_layout)
        
        # 按钮区域
        btn_layout = QHBoxLayout()
        self.send_btn = QPushButton("开始发送邮件")
        self.send_btn.setMinimumHeight(40)
        self.send_btn.setFont(font)
        self.send_btn.setStyleSheet("""
            QPushButton {
                background-color: #27AE60;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #219952;
            }
            QPushButton:pressed {
                background-color: #1E8449;
            }
        """)
        self.send_btn.clicked.connect(self.send_emails)
        btn_layout.addWidget(self.send_btn)
        
        # 添加测试连接按钮
        test_btn = QPushButton("测试邮箱连接")
        test_btn.setStyleSheet("""
            QPushButton {
                background-color: #3498DB;
            }
        """)
        test_btn.clicked.connect(self.test_outlook_connection)
        btn_layout.addWidget(test_btn)
        control_layout.addLayout(btn_layout)
        
        # 使用说明区域
        info_frame = QFrame()
        info_frame.setStyleSheet("""
            QFrame {
                background-color: #ECF0F1;
                border-radius: 5px;
                padding: 5px;
            }
        """)
        info_layout = QVBoxLayout(info_frame)
        
        # 添加标题
        help_title = QLabel("使用帮助")
        help_title.setStyleSheet("font-weight: bold; color: #2C3E50;")
        info_layout.addWidget(help_title)
        
        info_label = QLabel("1.选择Excel文件  2.选择收件人列  3.选择模板  4.点击发送")
        info_label.setWordWrap(True)
        info_layout.addWidget(info_label)
        
        # 注意事项
        note_label = QLabel("注意: 自动发送需要正常连接到Outlook，如无法连接将使用备用方式打开预览")
        note_label.setWordWrap(True)
        note_label.setStyleSheet("color: #E74C3C;")
        info_layout.addWidget(note_label)
        
        # 添加附件说明提示
        attachment_note = QLabel("如果指定了附件目录，将只在该目录中查找附件，否则将使用默认目录")
        attachment_note.setWordWrap(True)
        attachment_note.setStyleSheet("color: #555;")
        info_layout.addWidget(attachment_note)
        
        control_layout.addWidget(info_frame)
        
        # 状态显示
        status_layout = QHBoxLayout()
        status_layout.addWidget(QLabel("状态:"))
        self.status_label = QLabel("就绪")
        self.status_label.setStyleSheet("font-weight: bold; color: #2980B9;")
        status_layout.addWidget(self.status_label, 1)
        control_layout.addLayout(status_layout)
        
        control_group.setLayout(control_layout)
        send_layout.addWidget(control_group)
        
        send_tab.setLayout(send_layout)
        
        # 模板管理选项卡
        template_tab = QWidget()
        template_layout = QVBoxLayout()
        
        # 模板管理组
        template_group = QGroupBox("模板管理")
        template_group_layout = QHBoxLayout()
        
        # 模板列表
        template_list_layout = QVBoxLayout()
        template_list_layout.addWidget(QLabel("模板列表:"))
        self.template_list = QListWidget()
        self.template_list.setStyleSheet("border: 1px solid #CCC; border-radius: 4px;")
        self.template_list.currentRowChanged.connect(self.select_template)
        template_list_layout.addWidget(self.template_list)
        template_group_layout.addLayout(template_list_layout, 1)
        
        # 模板编辑
        template_edit_layout = QVBoxLayout()
        
        template_name_layout = QHBoxLayout()
        template_name_layout.addWidget(QLabel("模板名称:"))
        self.template_name = QLineEdit()
        template_name_layout.addWidget(self.template_name)
        template_edit_layout.addLayout(template_name_layout)
        
        template_subject_layout = QHBoxLayout()
        template_subject_layout.addWidget(QLabel("邮件主题:"))
        self.template_subject = QLineEdit()
        template_subject_layout.addWidget(self.template_subject)
        template_edit_layout.addLayout(template_subject_layout)
        
        template_edit_layout.addWidget(QLabel("邮件内容:"))
        self.template_content = QTextEdit()
        template_edit_layout.addWidget(self.template_content)
        
        template_btn_layout = QHBoxLayout()
        new_template_btn = QPushButton("新建模板")
        new_template_btn.setStyleSheet("""
            QPushButton {
                background-color: #3498DB;
            }
        """)
        new_template_btn.clicked.connect(self.new_template)
        template_btn_layout.addWidget(new_template_btn)
        
        save_template_btn = QPushButton("保存模板")
        save_template_btn.setStyleSheet("""
            QPushButton {
                background-color: #27AE60;
            }
            QPushButton:hover {
                background-color: #219952;
            }
        """)
        save_template_btn.clicked.connect(self.save_template)
        template_btn_layout.addWidget(save_template_btn)
        
        delete_template_btn = QPushButton("删除模板")
        delete_template_btn.setStyleSheet("""
            QPushButton {
                background-color: #E74C3C;
            }
            QPushButton:hover {
                background-color: #C0392B;
            }
        """)
        delete_template_btn.clicked.connect(self.delete_template)
        template_btn_layout.addWidget(delete_template_btn)
        template_edit_layout.addLayout(template_btn_layout)
        
        template_group_layout.addLayout(template_edit_layout, 2)
        template_group.setLayout(template_group_layout)
        
        template_layout.addWidget(template_group)
        
        # 添加模板使用说明
        template_help_frame = QFrame()
        template_help_frame.setStyleSheet("""
            QFrame {
                background-color: #ECF0F1;
                border-radius: 5px;
                padding: 10px;
                margin-top: 10px;
            }
        """)
        template_help_layout = QVBoxLayout(template_help_frame)
        
        template_help_title = QLabel("变量使用帮助")
        template_help_title.setStyleSheet("font-weight: bold; color: #2C3E50;")
        template_help_layout.addWidget(template_help_title)
        
        template_help_text = QLabel(
            "变量格式: {变量名}\n"
            "示例:\n"
            "尊敬的{姓名}先生/女士，您好！\n"
            "感谢您对{公司}的支持！"
        )
        template_help_text.setWordWrap(True)
        template_help_layout.addWidget(template_help_text)
        
        template_layout.addWidget(template_help_frame)
        
        template_tab.setLayout(template_layout)
        
        # 关于页面选项卡
        about_tab = QWidget()
        about_layout = QVBoxLayout()
        
        # 使用滚动区域以适应内容大小
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_content = QWidget()
        scroll_layout = QVBoxLayout(scroll_content)
        
        # 添加标题
        title_label = QLabel("邮件群发助手")
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet("font-size: 16px; font-weight: bold; color: #2980B9; margin: 5px;")
        scroll_layout.addWidget(title_label)
        
        # 添加版本信息
        version_label = QLabel("版本: 1.0.0")
        version_label.setAlignment(Qt.AlignCenter)
        version_label.setStyleSheet("font-size: 11px; color: #555; margin-bottom: 10px;")
        scroll_layout.addWidget(version_label)
        
        # 添加软件描述
        desc_frame = QFrame()
        desc_frame.setStyleSheet("background-color: white; border: 1px solid #DDD; border-radius: 5px; padding: 12px;")
        desc_layout = QVBoxLayout(desc_frame)
        desc_layout.setContentsMargins(8, 8, 8, 8)
        
        # 添加一个标签来说明这是README内容
        readme_title = QLabel("软件说明")
        readme_title.setStyleSheet("font-size: 13px; font-weight: bold; color: #333; margin-bottom: 4px;")
        desc_layout.addWidget(readme_title)
        
        # 使用纯文本方式显示README内容
        desc_text = QTextEdit()
        try:
            readme_path = resource_path("README.md")
            with open(readme_path, "r", encoding="utf-8") as f:
                readme_content = f.read()
                
            # 移除Markdown标记，转为纯文本
            readme_text = readme_content
            # 替换标题标记
            for i in range(5, 0, -1):  # 从##### 到 #
                heading_mark = '#' * i + ' '
                readme_text = readme_text.replace(heading_mark, '')
            
            # 处理列表项
            lines = readme_text.split('\n')
            for i, line in enumerate(lines):
                if line.strip().startswith('- '):
                    # 保留列表符号但替换为普通文本的列表形式
                    lines[i] = "• " + line.strip()[2:]
            readme_text = '\n'.join(lines)
            
            desc_text.setPlainText(readme_text)
        except Exception as e:
            desc_text.setPlainText(f"无法加载README文件: {str(e)}\n文件路径: {readme_path}")
            traceback.print_exc()  # 打印完整错误信息
        
        # 设置字体和样式
        font = QFont("Microsoft YaHei", 10)  # 增大字体
        desc_text.setFont(font)
        desc_text.setReadOnly(True)
        desc_text.setStyleSheet("""
            QTextEdit {
                border: none;
                background-color: white;
                font-size: 10pt;
                line-height: 140%;
            }
        """)
        
        # 调整行间距和文本边距
        desc_text.document().setDocumentMargin(5)
        
        # 设置高度，避免过大
        desc_text.setMinimumHeight(280)
        desc_text.setMaximumHeight(350)
        desc_layout.addWidget(desc_text)
        
        scroll_layout.addWidget(desc_frame)
        
        # 添加作者和打赏信息
        author_frame = QFrame()
        author_frame.setStyleSheet("background-color: white; border: 1px solid #DDD; border-radius: 5px; padding: 10px; margin-top: 15px;")
        author_layout = QVBoxLayout(author_frame)
        author_layout.setContentsMargins(8, 8, 8, 8)
        
        # 标题
        support_title = QLabel("支持与赞赏")
        support_title.setAlignment(Qt.AlignCenter)
        support_title.setStyleSheet("font-size: 13px; font-weight: bold; color: #E74C3C; margin-bottom: 6px;")
        author_layout.addWidget(support_title)
        
        # 添加打赏图片（如果存在）
        tip_image_container = QHBoxLayout()
        
        # 检查是否存在打赏图片
        donation_images = []
        wechat_path = resource_path("donation_wechat.jpg")
        alipay_path = resource_path("donation_alipay.jpg")
        
        if os.path.exists(wechat_path):
            donation_images.append(("微信赞赏", wechat_path))
        if os.path.exists(alipay_path):
            donation_images.append(("支付宝赞赏", alipay_path))
        
        # 如果找到打赏图片，则添加
        if donation_images:
            for title, image_path in donation_images:
                image_layout = QVBoxLayout()
                image_label = QLabel()
                pixmap = QPixmap(image_path)
                scaled_pixmap = pixmap.scaled(140, 140, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                image_label.setPixmap(scaled_pixmap)
                image_label.setAlignment(Qt.AlignCenter)
                
                title_label = QLabel(title)
                title_label.setAlignment(Qt.AlignCenter)
                title_label.setStyleSheet("font-weight: bold; margin-top: 3px; font-size: 11px;")
                
                image_layout.addWidget(image_label)
                image_layout.addWidget(title_label)
                tip_image_container.addLayout(image_layout)
        else:
            # 如果没有打赏图片，则显示提示
            no_image_label = QLabel("如需添加打赏码，请将支付码图片\n命名为donation_wechat.jpg或donation_alipay.jpg\n放在程序根目录下")
            no_image_label.setAlignment(Qt.AlignCenter)
            no_image_label.setStyleSheet("color: #555; margin: 12px; font-size: 11px;")
            tip_image_container.addWidget(no_image_label)
            
        author_layout.addLayout(tip_image_container)
        
        # 联系信息
        contact_label = QLabel("邮箱: help@kiramao.cn")
        contact_label.setAlignment(Qt.AlignCenter)
        contact_label.setStyleSheet("color: #555; margin-top: 6px; font-size: 11px;")
        author_layout.addWidget(contact_label)
        
        scroll_layout.addWidget(author_frame)
        
        # 设置内容到滚动区域
        scroll_area.setWidget(scroll_content)
        about_layout.addWidget(scroll_area)
        
        about_tab.setLayout(about_layout)
        
        # 添加选项卡到主窗口，并调整字体大小
        tab_font = QFont("Microsoft YaHei", 10)  # 使用微软雅黑字体，大小设为10
        
        # 发送邮件选项卡
        tabs.addTab(send_tab, "发送邮件")
        tabs.setTabText(0, "发送")
        
        # 模板管理选项卡
        tabs.addTab(template_tab, "模板管理")
        tabs.setTabText(1, "模板")
        
        # 关于软件选项卡
        tabs.addTab(about_tab, "关于")
        tabs.setTabText(2, "关于")
        
        # 设置所有选项卡的字体
        for i in range(tabs.count()):
            tabs.tabBar().setTabTextColor(i, QColor("#444444"))
            tabs.tabBar().setFont(tab_font)
        
        # 调整选项卡控件的样式
        tabs.setStyleSheet("""
            QTabBar::tab {
                padding: 6px 12px;
                margin-right: 2px;
                min-width: 60px;
                max-width: 90px;
                font-weight: bold;
            }
        """)
        
        main_layout.addWidget(tabs)
        main_widget.setLayout(main_layout)
        self.setCentralWidget(main_widget)
        
        # 初始化数据
        self.refresh_template_list()
    
    def detect_email_clients(self):
        """检测可用的邮件客户端"""
        try:
            self.status_label.setText("正在检测可用的邮件客户端...")
            QApplication.processEvents()
            
            # 记住当前选择的客户端
            current_client = self.client_combo.currentData() if self.client_combo.currentIndex() >= 0 else None
            
            # 清空并重新填充客户端下拉框
            self.client_combo.clear()
            
            # 获取可用的客户端
            clients = self.outlook_sender.get_available_clients()
            
            if clients:
                for client_id, client_name in clients.items():
                    self.client_combo.addItem(client_name, client_id)
                    
                # 如果之前有选择，尝试恢复选择
                if current_client:
                    for i in range(self.client_combo.count()):
                        if self.client_combo.itemData(i) == current_client:
                            self.client_combo.setCurrentIndex(i)
                            break
                            
                self.status_label.setText(f"已检测到 {len(clients)} 个邮件客户端")
                
                # 将客户端选择连接到更改处理函数
                self.client_combo.currentIndexChanged.connect(self.client_changed)
            else:
                self.status_label.setText("未检测到可用的邮件客户端")
        except Exception as e:
            self.status_label.setText(f"检测邮件客户端失败: {str(e)}")
            print(traceback.format_exc())
    
    def client_changed(self, index):
        """处理邮件客户端选择变更"""
        if index >= 0:
            client_id = self.client_combo.itemData(index)
            if client_id:
                self.outlook_sender.set_client(client_id)
                self.load_sender_accounts()  # 重新加载当前客户端的账户
    
    def load_sender_accounts(self):
        """加载发件人邮箱账户列表"""
        try:
            self.status_label.setText("正在连接邮箱账户...")
            QApplication.processEvents()
            
            # 如果还没有初始化客户端选择框，先检测客户端
            if self.client_combo.count() == 0:
                self.detect_email_clients()
            
            # 清空发件人下拉框
            self.sender_combo.clear()
            self.sender_combo.addItem("自动选择")
            
            # 获取当前选择的邮件客户端
            client_id = self.client_combo.currentData() if self.client_combo.currentIndex() >= 0 else None
            
            # 只有当选择Outlook时才获取账户列表
            if client_id == self.outlook_sender.CLIENT_OUTLOOK:
                # 获取Outlook账户列表
                profiles = self.outlook_sender.get_sender_profiles()
                
                if profiles:
                    for profile in profiles:
                        self.sender_combo.addItem(f"{profile['name']} <{profile['email']}>", profile['email'])
                    self.status_label.setText(f"已找到 {len(profiles)} 个邮箱账户")
                else:
                    self.status_label.setText("未找到邮箱账户，将使用备用方法")
            else:
                # 其他客户端不支持获取账户列表
                self.status_label.setText(f"已选择 {self.client_combo.currentText()} 作为邮件客户端")
        except Exception as e:
            self.status_label.setText(f"加载邮箱账户失败: {str(e)}")
            print(traceback.format_exc())
    
    def test_outlook_connection(self):
        """测试Outlook连接"""
        try:
            # 获取当前选择的邮件客户端
            client_id = self.client_combo.currentData() if self.client_combo.currentIndex() >= 0 else None
            
            # 如果不是Outlook，显示提示
            if client_id != self.outlook_sender.CLIENT_OUTLOOK:
                QMessageBox.information(self, "邮件客户端",
                                       f"当前选择的邮件客户端是: {self.client_combo.currentText()}\n\n"
                                       f"只有在使用Outlook时才能测试连接和获取账户列表。")
                return
            
            self.status_label.setText("正在测试Outlook连接...")
            QApplication.processEvents()
            
            if self.outlook_sender.connect_outlook():
                profiles = self.outlook_sender.get_sender_profiles()
                if profiles:
                    account_names = [f"{p['name']} <{p['email']}>" for p in profiles]
                    accounts_text = "\n".join(account_names)
                    QMessageBox.information(self, "连接成功", 
                                           f"成功连接到Outlook！\n\n找到以下账户:\n{accounts_text}\n\n您可以使用自动发送功能直接发送邮件。")
                    self.status_label.setText("Outlook连接测试成功")
                else:
                    QMessageBox.warning(self, "部分成功", 
                                      "连接到Outlook成功，但未找到邮箱账户。\n可能需要先打开Outlook并登录您的账户。")
                    self.status_label.setText("Outlook连接部分成功")
            else:
                if self.auto_send_checkbox.isChecked():
                    QMessageBox.critical(self, "连接失败 - 自动发送不可用", 
                                       "无法连接到Outlook，自动发送功能将不可用。\n\n将使用备用方法创建邮件预览。")
                else:
                    QMessageBox.critical(self, "连接失败", 
                                       "无法连接到Outlook。\n可能的原因:\n- Outlook未安装或未运行\n- 权限不足\n- COM组件注册问题\n\n将使用备用方法创建邮件预览。")
                self.status_label.setText("Outlook连接测试失败")
        except Exception as e:
            QMessageBox.critical(self, "测试错误", f"测试Outlook连接时发生错误:\n{str(e)}")
            self.status_label.setText("Outlook连接测试出错")
            print(traceback.format_exc())
    
    def browse_excel(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "选择Excel文件", "", "Excel文件 (*.xlsx *.xls)")
        if file_path:
            self.excel_path.setText(file_path)
            # 加载Sheet列表
            sheets = self.excel_reader.get_sheet_names(file_path)
            self.sheet_combo.clear()
            self.sheet_combo.addItems(sheets)
    
    def load_excel_data(self):
        if not self.excel_path.text():
            QMessageBox.warning(self, "警告", "请先选择Excel文件")
            return
            
        sheet_name = self.sheet_combo.currentText()
        if not sheet_name:
            QMessageBox.warning(self, "警告", "请选择Sheet页")
            return
        
        try:
            self.status_label.setText("正在加载Excel数据...")
            QApplication.processEvents()
            
            # 加载列名
            columns = self.excel_reader.get_column_names(self.excel_path.text(), sheet_name)
            
            # 设置收件人邮箱列下拉框
            self.email_column_combo.clear()
            self.email_column_combo.addItems(columns)
            
            # 设置默认选择邮箱列
            for i, col in enumerate(columns):
                if "邮箱" in col or "email" in col.lower() or "mail" in col.lower():
                    self.email_column_combo.setCurrentIndex(i)
                    break
            
            # 加载变量
            self.var_list.clear()
            for col in columns:
                self.var_list.addItem(f"{{{col}}}")
                
            self.status_label.setText(f"Excel数据加载完成，共 {len(columns)} 列")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"加载Excel数据出错: {str(e)}")
            self.status_label.setText("Excel数据加载失败")
            print(traceback.format_exc())
    
    def refresh_template_list(self):
        templates = self.template_manager.get_templates()
        self.template_list.clear()
        self.template_combo.clear()
        for template in templates:
            self.template_list.addItem(template)
            self.template_combo.addItem(template)
    
    def load_template(self):
        template_name = self.template_combo.currentText()
        if template_name:
            subject, content = self.template_manager.get_template_content(template_name)
            self.mail_subject.setText(subject)
            self.mail_content.setText(content)
    
    def select_template(self):
        template_name = self.template_list.currentItem().text() if self.template_list.currentItem() else ""
        if template_name:
            subject, content = self.template_manager.get_template_content(template_name)
            self.template_name.setText(template_name)
            self.template_subject.setText(subject)
            self.template_content.setText(content)
    
    def new_template(self):
        self.template_name.clear()
        self.template_subject.clear()
        self.template_content.clear()
    
    def save_template(self):
        name = self.template_name.text()
        subject = self.template_subject.text()
        content = self.template_content.toPlainText()
        
        if not name:
            QMessageBox.warning(self, "警告", "请输入模板名称")
            return
        
        self.template_manager.save_template(name, subject, content)
        self.refresh_template_list()
        QMessageBox.information(self, "成功", "模板保存成功")
    
    def delete_template(self):
        template_name = self.template_list.currentItem().text() if self.template_list.currentItem() else ""
        if not template_name:
            QMessageBox.warning(self, "警告", "请选择要删除的模板")
            return
        
        reply = QMessageBox.question(self, "确认", f"是否确认删除模板'{template_name}'?", 
                                    QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.template_manager.delete_template(template_name)
            self.refresh_template_list()
            QMessageBox.information(self, "成功", "模板删除成功")
    
    def browse_attachment_dir(self):
        dir_path = QFileDialog.getExistingDirectory(self, "选择附件所在目录")
        if dir_path:
            self.attachment_dir.setText(dir_path)
    
    def send_emails(self):
        if not self.excel_path.text():
            QMessageBox.warning(self, "警告", "请选择Excel文件")
            return
            
        sheet_name = self.sheet_combo.currentText()
        if not sheet_name:
            QMessageBox.warning(self, "警告", "请选择Sheet页")
            return
        
        to_column = self.email_column_combo.currentText()
        if not to_column:
            QMessageBox.warning(self, "警告", "请选择收件人邮箱列")
            return
        
        subject = self.mail_subject.text()
        content = self.mail_content.toPlainText()
        
        if not subject or not content:
            QMessageBox.warning(self, "警告", "邮件主题和内容不能为空")
            return
        
        # 获取选择的发件人邮箱
        sender_email = None
        sender_idx = self.sender_combo.currentIndex()
        if sender_idx > 0:  # 不是"自动选择"
            sender_email = self.sender_combo.itemData(sender_idx)
        
        # 获取自动发送选项
        auto_send = self.auto_send_checkbox.isChecked()
        
        # 获取附件模式和目录
        attachment_pattern = self.attachment_pattern.text().strip()
        attachment_dir = self.attachment_dir.text().strip()
        
        try:
            self.status_label.setText("正在读取Excel数据...")
            QApplication.processEvents()
            
            data = self.excel_reader.read_data(self.excel_path.text(), sheet_name)
            if not data:
                QMessageBox.warning(self, "警告", "Excel文件中没有数据")
                self.status_label.setText("没有找到数据")
                return
            
            # 检查是否开启自动发送但无法连接Outlook
            current_client = self.client_combo.currentData()
            if auto_send and current_client == self.outlook_sender.CLIENT_OUTLOOK and not self.outlook_sender.connect_outlook():
                confirm = QMessageBox.question(self, "无法自动发送", 
                                            "已开启自动发送模式，但无法连接到Outlook。\n\n是否继续使用预览模式发送邮件？",
                                            QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
                if confirm == QMessageBox.No:
                    self.status_label.setText("操作已取消")
                    return
                auto_send = False
            
            # 非Outlook客户端不支持自动发送
            if auto_send and current_client != self.outlook_sender.CLIENT_OUTLOOK:
                QMessageBox.information(self, "功能限制", 
                                       f"当前选择的邮件客户端 ({self.client_combo.currentText()}) 不支持自动发送功能。\n\n"
                                       f"将使用预览模式创建邮件。")
                auto_send = False
            
            # 确认发送
            if auto_send:
                confirm_text = f"将向{len(data)}个收件人直接发送邮件（无预览），是否继续?"
            else:
                confirm_text = f"将为{len(data)}个收件人创建邮件预览窗口，是否继续?"
                
            confirm = QMessageBox.question(self, "确认", confirm_text,
                                        QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if confirm == QMessageBox.Yes:
                try:
                    self.send_btn.setEnabled(False)
                    self.status_label.setText("正在创建邮件...")
                    QApplication.processEvents()
                    
                    sent_count = self.outlook_sender.send_batch_emails(data, to_column, subject, content, sender_email, auto_send, attachment_pattern, attachment_dir)
                    
                    if sent_count > 0:
                        if auto_send:
                            self.status_label.setText(f"已成功发送 {sent_count} 封邮件")
                            QMessageBox.information(self, "成功", f"已成功发送{sent_count}封邮件。")
                        else:
                            self.status_label.setText(f"已成功创建 {sent_count} 封邮件")
                            QMessageBox.information(self, "成功", f"已成功创建{sent_count}封邮件。\n\n如果您使用的是Outlook，请在Outlook中检查并发送这些邮件。\n如果使用其他邮件客户端，这些邮件已经在默认邮件程序中打开。")
                    else:
                        self.status_label.setText("没有创建任何邮件")
                        QMessageBox.warning(self, "警告", "没有创建任何邮件，请检查数据和邮件客户端配置。")
                except Exception as e:
                    error_msg = str(e)
                    detailed_msg = ""
                    
                    if "无法连接到Outlook" in error_msg:
                        detailed_msg = (
                            "1. 确认Outlook已安装并能正常运行\n"
                            "2. 尝试手动启动Outlook，然后再次运行此程序\n"
                            "3. 如果问题依然存在，请尝试以管理员身份运行此程序"
                        )
                    
                    self.status_label.setText(f"发送邮件出错: {error_msg}")
                    QMessageBox.critical(self, "错误", f"发送邮件时出错: {error_msg}\n\n{detailed_msg}")
                finally:
                    self.send_btn.setEnabled(True)
        except Exception as e:
            self.status_label.setText(f"处理Excel数据出错: {str(e)}")
            QMessageBox.critical(self, "错误", f"处理Excel数据时出错: {str(e)}")
            print(traceback.format_exc()) 