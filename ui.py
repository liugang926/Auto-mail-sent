import os
import sys
from PyQt5.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                            QLabel, QPushButton, QLineEdit, QFileDialog, 
                            QSpinBox, QTextEdit, QProgressBar, QComboBox,
                            QGroupBox, QFormLayout, QMessageBox, QDialog,
                            QListWidget)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QTimer
from PyQt5.QtGui import QFont, QPixmap, QIcon
from qt_material import apply_stylesheet
from email_processor import EmailSender
from word_reader import WordReader
from excel_reader import ExcelReader
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header
from datetime import datetime
from PyQt5.QtWidgets import QApplication

class BlurredWidget(QWidget):
    """实现毛玻璃效果的基础Widget"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setStyleSheet("""
            QWidget {
                background-color: rgba(255, 255, 255, 180);
                border-radius: 10px;
            }
        """)
        
class EmailPreviewWidget(QWidget):
    """修改邮件预览窗口以支持HTML格式"""
    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        
        self.subject_label = QLabel("主题: ")
        self.to_label = QLabel("收件人: ")
        self.content = QTextEdit()
        self.content.setReadOnly(True)
        
        # 设置预览窗口的样式
        self.content.setStyleSheet("""
            QTextEdit {
                background-color: white;
                border: 1px solid #ddd;
                border-radius: 4px;
                padding: 10px;
            }
        """)
        
        layout.addWidget(self.subject_label)
        layout.addWidget(self.to_label)
        layout.addWidget(self.content)
        
    def update_preview(self, subject, to_name, to_email, content):
        self.subject_label.setText(f"主题: {subject}")
        self.to_label.setText(f"收件人: {to_name} <{to_email}>")
        # 直接设置HTML内容
        self.content.setHtml(content)

class EmailSenderThread(QThread):
    """邮件发送线程"""
    progress_updated = pyqtSignal(int)
    sending_finished = pyqtSignal()
    error_occurred = pyqtSignal(str)
    
    def __init__(self, email_sender, excel_data, template_content, subject,
                 name_column, email_column, interval):
        super().__init__()
        self.email_sender = email_sender
        self.excel_data = excel_data
        self.template_content = template_content
        self.subject = subject
        self.name_column = name_column
        self.email_column = email_column
        self.interval = interval
        self.is_running = True
    
    def run(self):
        try:
            total = len(self.excel_data)
            for i, row in enumerate(self.excel_data):
                if not self.is_running:
                    break
                    
                # 替换模板中的变量
                content = self.template_content
                for col, value in row.items():
                    content = content.replace(f"{{{col}}}", str(value))
                
                # 发送邮件
                self.email_sender.send_email(
                    row[self.email_column],
                    self.subject,
                    content
                )
                
                # 更新进度
                progress = int((i + 1) / total * 100)
                self.progress_updated.emit(progress)
                
                # 等待指定时间
                if i < total - 1:  # 最后一封邮件不需要等待
                    QThread.sleep(self.interval)
            
            self.sending_finished.emit()
            
        except Exception as e:
            self.error_occurred.emit(str(e))
    
    def stop(self):
        self.is_running = False

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("智能邮件群发系统")
        self.setMinimumSize(900, 700)
        
        # 设置窗口图标
        icon = QIcon(":/icons/email.png")  # 使用Qt资源系统
        self.setWindowIcon(icon)
        
        # 设置默认字体
        app = QApplication.instance()
        font = QFont("Microsoft YaHei UI", 9)  # 使用微软雅黑UI字体
        app.setFont(font)
        
        # 设置窗口背景
        self.setObjectName("mainWindow")
        
        # 初始化读取器和发送器
        self.word_reader = WordReader()
        self.excel_reader = ExcelReader()
        self.email_sender = EmailSender()
        
        # 数据存储
        self.template_content = ""
        self.excel_data = None
        self.name_column = ""
        self.email_column = ""
        
        # 先创建UI
        self.setup_ui()
        
        # 设置特殊按钮的ObjectName (移到UI创建之后)
        self.test_send_btn.setObjectName("test_send_btn")
        self.stop_btn.setObjectName("stop_btn")
        
        # 应用样式
        self.apply_blur_style()
    
    def setup_ui(self):
        """修改UI设置，移除不需要的组件"""
        # 主容器
        central_widget = QWidget()
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)
        
        # ===== 文件选择区域 =====
        file_group = QGroupBox()
        file_group.setTitle("文件选择")
        file_layout = QFormLayout()
        file_layout.setSpacing(12)
        file_layout.setContentsMargins(15, 25, 15, 15)
        
        # Word模板选择
        word_layout = QHBoxLayout()
        self.word_path = QLineEdit()
        self.word_path.setReadOnly(True)
        self.word_path.setMinimumHeight(32)  # 增加高度
        word_browse_btn = QPushButton("浏览...")
        word_browse_btn.setFixedSize(90, 32)  # 固定按钮大小
        word_browse_btn.clicked.connect(self.browse_word)
        word_layout.addWidget(self.word_path)
        word_layout.addWidget(word_browse_btn)
        word_layout.setSpacing(10)
        
        # Excel数据选择
        excel_layout = QHBoxLayout()
        self.excel_path = QLineEdit()
        self.excel_path.setReadOnly(True)
        self.excel_path.setMinimumHeight(32)  # 增加高度
        excel_browse_btn = QPushButton("浏览...")
        excel_browse_btn.setFixedSize(90, 32)  # 固定按钮大小
        excel_browse_btn.clicked.connect(self.browse_excel)
        excel_layout.addWidget(self.excel_path)
        excel_layout.addWidget(excel_browse_btn)
        excel_layout.setSpacing(10)
        
        file_layout.addRow("Word模板:", word_layout)
        file_layout.addRow("Excel数据:", excel_layout)
        file_group.setLayout(file_layout)
        
        # ===== 邮件配置区域 =====
        config_group = QGroupBox("邮件配置")
        config_layout = QFormLayout()
        config_layout.setSpacing(12)
        config_layout.setContentsMargins(15, 25, 15, 15)
        
        # 变量匹配状态显示
        self.variables_status = QLabel("变量匹配状态")
        self.variables_status.setWordWrap(True)
        
        # 主题输入框
        self.subject_input = QLineEdit()
        self.subject_input.setMinimumHeight(32)
        
        # 发送间隔设置
        self.interval_spinbox = QSpinBox()
        self.interval_spinbox.setMinimumHeight(32)
        self.interval_spinbox.setRange(1, 600)
        self.interval_spinbox.setValue(30)
        self.interval_spinbox.setSuffix(" 秒")
        
        # 将组件添加到配置布局
        config_layout.addRow("变量状态:", self.variables_status)
        config_layout.addRow("邮件主题:", self.subject_input)
        config_layout.addRow("发送间隔:", self.interval_spinbox)
        
        # 测试按钮和帮助按钮
        test_btn_layout = QHBoxLayout()
        self.test_send_btn = QPushButton("测试邮箱配置")
        self.test_send_btn.setFixedSize(120, 36)
        self.test_send_btn.clicked.connect(self.test_email_config)

        # 添加帮助按钮
        help_btn = QPushButton("帮助")
        help_btn.setObjectName("help_btn")
        help_btn.setFixedSize(80, 36)
        help_btn.clicked.connect(self.show_help)

        test_btn_layout.addWidget(self.test_send_btn)
        test_btn_layout.addWidget(help_btn)
        test_btn_layout.addStretch()
        config_layout.addRow("", test_btn_layout)
        
        config_group.setLayout(config_layout)
        
        # ===== 预览和进度区域 =====
        bottom_layout = QHBoxLayout()
        bottom_layout.setSpacing(15)
        
        # 预览区域
        preview_group = QGroupBox()
        preview_group.setTitle("邮件预览")
        preview_layout = QVBoxLayout()
        preview_layout.setSpacing(12)
        preview_layout.setContentsMargins(15, 25, 15, 15)
        
        self.preview_widget = EmailPreviewWidget()
        preview_layout.addWidget(self.preview_widget)
        
        # 进度区域
        progress_group = QGroupBox()
        progress_group.setTitle("发送进度")
        progress_layout = QVBoxLayout()
        progress_layout.setSpacing(12)
        progress_layout.setContentsMargins(15, 25, 15, 15)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimumHeight(24)
        self.status_label = QLabel("就绪")
        self.status_label.setMinimumHeight(36)
        
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(10)
        
        self.send_btn = QPushButton("开始发送")
        self.send_btn.setFixedHeight(36)
        self.send_btn.clicked.connect(self.start_sending)
        
        self.stop_btn = QPushButton("停止发送")
        self.stop_btn.setFixedHeight(36)
        self.stop_btn.clicked.connect(self.stop_sending)
        self.stop_btn.setEnabled(False)
        
        btn_layout.addWidget(self.send_btn)
        btn_layout.addWidget(self.stop_btn)
        
        progress_layout.addWidget(self.progress_bar)
        progress_layout.addWidget(self.status_label)
        progress_layout.addLayout(btn_layout)
        progress_layout.addStretch()
        
        progress_group.setLayout(progress_layout)
        
        # 设置预览和进度区域的比例
        bottom_layout.addWidget(preview_group, 2)
        bottom_layout.addWidget(progress_group, 1)
        
        # 添加所有组件到主布局
        main_layout.addWidget(file_group)
        main_layout.addWidget(config_group)
        main_layout.addLayout(bottom_layout, 1)
        
        self.setCentralWidget(central_widget)
    
    def apply_blur_style(self):
        """应用现代化UI风格"""
        self.setStyleSheet("""
            * {
                font-family: "Microsoft YaHei UI", "Microsoft YaHei", "SimHei", sans-serif;
            }
            QMainWindow {
                background-color: #f8f9fa;
            }
            QGroupBox {
                background-color: white;
                border-radius: 8px;
                border: 1px solid #e9ecef;
                margin-top: 20px;
                padding: 28px 15px 15px 15px;
                font-weight: 500;
                font-size: 14px;
                color: #2c3e50;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                left: 15px;
                top: 10px;
                padding: 0px 10px;
                background-color: white;
                color: #2c3e50;
                font-size: 14px;
                font-weight: 500;
            }
            QPushButton {
                background-color: #3498db;
                color: white;
                border-radius: 4px;
                padding: 8px 16px;
                border: none;
                font-weight: 500;
                font-size: 13px;
                min-width: 80px;
                min-height: 32px;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
            QPushButton:pressed {
                background-color: #2473a7;
            }
            QPushButton:disabled {
                background-color: #bdc3c7;
            }
            QLineEdit, QTextEdit, QComboBox, QSpinBox {
                background-color: white;
                border-radius: 4px;
                border: 1px solid #ced4da;
                padding: 6px 12px;
                color: #2c3e50;
                font-size: 13px;
                min-height: 32px;
            }
            QLineEdit:focus, QTextEdit:focus, QComboBox:focus, QSpinBox:focus {
                border: 2px solid #3498db;
                background-color: white;
            }
            QLabel {
                color: #2c3e50;
                font-size: 13px;
                padding: 4px 0;
                font-weight: normal;
            }
            QProgressBar {
                border: none;
                border-radius: 4px;
                text-align: center;
                background-color: #e9ecef;
                font-size: 12px;
                color: white;
                min-height: 24px;
            }
            QProgressBar::chunk {
                background-color: #2ecc71;
                border-radius: 4px;
            }
            
            /* 特殊按钮样式 */
            QPushButton#test_send_btn {
                background-color: #2ecc71;
            }
            QPushButton#test_send_btn:hover {
                background-color: #27ae60;
            }
            QPushButton#stop_btn {
                background-color: #e74c3c;
            }
            QPushButton#stop_btn:hover {
                background-color: #c0392b;
            }
            
            /* 下拉框样式 */
            QComboBox::drop-down {
                border: none;
                width: 30px;
            }
            QComboBox::down-arrow {
                image: none;
                border-left: 5px solid transparent;
                border-right: 5px solid transparent;
                border-top: 5px solid #495057;
                margin-right: 8px;
            }
            
            /* 帮助按钮样式 */
            QPushButton#help_btn {
                background-color: #6c757d;
                color: white;
                border-radius: 4px;
                padding: 8px 16px;
                border: none;
                font-weight: bold;
                font-size: 13px;
            }
            QPushButton#help_btn:hover {
                background-color: #5a6268;
            }
            QPushButton#help_btn:pressed {
                background-color: #545b62;
            }
        """)
    
    def browse_word(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择Word模板", "", "Word文档 (*.docx *.doc)"
        )
        if file_path:
            self.word_path.setText(file_path)
            try:
                self.template_content, self.template_variables = self.word_reader.read_template(file_path)
                # 显示找到的变量
                variables_text = "模板中的变量：\n" + "\n".join([f"{{{var}}}" for var in self.template_variables])
                self.variables_status.setText(variables_text)
                
                # 如果已经加载了Excel，检查变量匹配
                if self.excel_data:
                    self.check_variable_matching()
                
                QMessageBox.information(self, "成功", f"Word模板加载成功!\n找到 {len(self.template_variables)} 个变量。")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"无法读取Word文档: {str(e)}")
    
    def browse_excel(self):
        """修改Excel文件选择处理"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择Excel数据文件", "", "Excel文件 (*.xlsx *.xls)"
        )
        if file_path:
            self.excel_path.setText(file_path)
            try:
                self.excel_data, self.excel_columns = self.excel_reader.read_data(file_path)
                
                # 如果已经加载了Word模板，检查变量匹配
                if hasattr(self, 'template_variables'):
                    self.check_variable_matching()
                
                QMessageBox.information(self, "成功", f"Excel数据加载成功！共{len(self.excel_data)}条记录。")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"无法读取Excel文件: {str(e)}")
    
    def check_variable_matching(self):
        """检查Word模板变量与Excel列的匹配情况，并自动生成预览"""
        if not hasattr(self, 'template_variables') or not hasattr(self, 'excel_columns'):
            return
        
        # 检查变量匹配
        matched_vars = []
        unmatched_vars = []
        self.name_column = None
        self.email_column = None
        
        # 自动识别姓名和邮箱列
        for col in self.excel_columns:
            if not self.name_column and ("姓名" in col or "名字" in col or "name" in col.lower()):
                self.name_column = col
            if not self.email_column and ("邮箱" in col or "邮件" in col or "email" in col.lower()):
                self.email_column = col
        
        # 检查其他变量匹配
        for var in self.template_variables:
            if var in self.excel_columns:
                matched_vars.append(var)
            else:
                unmatched_vars.append(var)
        
        # 更新变量状态显示
        status_text = "变量匹配状态：\n\n"
        
        # 显示姓名和邮箱列匹配状态
        if self.name_column:
            status_text += f"✅ 姓名列: {self.name_column}\n"
        else:
            status_text += "❌ 未找到姓名列\n"
        
        if self.email_column:
            status_text += f"✅ 邮箱列: {self.email_column}\n"
        else:
            status_text += "❌ 未找到邮箱列\n"
        
        status_text += "\n其他变量匹配：\n"
        if matched_vars:
            status_text += "✅ 已匹配变量：\n" + "\n".join([f"{{{var}}}" for var in matched_vars]) + "\n\n"
        if unmatched_vars:
            status_text += "❌ 未匹配变量：\n" + "\n".join([f"{{{var}}}" for var in unmatched_vars])
        
        self.variables_status.setText(status_text)
        
        # 自动生成预览
        self.auto_generate_preview(unmatched_vars)
        
        # 显示警告信息
        warnings = []
        if not self.name_column:
            warnings.append("未找到姓名列")
        if not self.email_column:
            warnings.append("未找到邮箱列")
        if unmatched_vars:
            warnings.append(f"以下变量未找到对应列：{', '.join(unmatched_vars)}")
        
        if warnings:
            QMessageBox.warning(self, "警告", "\n".join(warnings))

    def auto_generate_preview(self, unmatched_vars=None):
        """自动生成预览"""
        if not self.template_content or not self.excel_data:
            return
            
        if not self.name_column or not self.email_column:
            return
            
        # 获取第一条数据作为预览
        try:
            first_row = self.excel_data[0]
            content = self.template_content
            
            # 替换所有匹配的变量
            for var in self.template_variables:
                if var in first_row:
                    content = content.replace(f"{{{var}}}", str(first_row[var]))
                elif var in unmatched_vars:
                    # 对于未匹配的变量，保留原样显示
                    content = content.replace(f"{{{var}}}", f"[未匹配变量: {{{var}}}]")
            
            # 获取主题（如果未输入，使用默认值）
            subject = self.subject_input.text() or "[请输入邮件主题]"
            
            # 更新预览
            self.preview_widget.update_preview(
                subject,
                first_row[self.name_column],
                first_row[self.email_column],
                content
            )
            
        except Exception as e:
            self.preview_widget.update_preview(
                "[请输入邮件主题]",
                "预览生成失败",
                "预览生成失败",
                f"生成预览时发生错误: {str(e)}"
            )
    
    def start_sending(self):
        """修改发送逻辑，移除变量选择相关代码"""
        if not self.template_content:
            QMessageBox.warning(self, "警告", "请先加载Word模板!")
            return
            
        if not self.excel_data:
            QMessageBox.warning(self, "警告", "请先加载Excel数据!")
            return
            
        if not self.name_column or not self.email_column:
            QMessageBox.warning(self, "警告", "未找到姓名或邮箱列!")
            return
            
        subject = self.subject_input.text()
        if not subject:
            QMessageBox.warning(self, "警告", "请输入邮件主题!")
            return
        
        # 创建发送线程
        self.sender_thread = EmailSenderThread(
            self.email_sender,
            self.excel_data,
            self.template_content,
            subject,
            self.name_column,
            self.email_column,
            self.interval_spinbox.value()
        )
        
        # 连接信号
        self.sender_thread.progress_updated.connect(self.update_progress)
        self.sender_thread.sending_finished.connect(self.sending_finished)
        self.sender_thread.error_occurred.connect(self.handle_sending_error)
        
        # 更新UI状态
        self.send_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.progress_bar.setValue(0)
        
        # 开始发送
        self.sender_thread.start()
    
    def stop_sending(self):
        if hasattr(self, "sender_thread") and self.sender_thread.isRunning():
            self.sender_thread.stop()
            self.status_label.setText("正在停止...")
            self.stop_btn.setEnabled(False)
    
    def update_progress(self, value):
        self.progress_bar.setValue(value)
    
    def sending_finished(self):
        self.send_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.status_label.setText("发送完成!")
        QMessageBox.information(self, "成功", "所有邮件已发送完成!")
    
    def handle_sending_error(self, error_msg):
        self.status_label.setText(f"错误: {error_msg}")
    
    def test_email_config(self):
        """测试邮箱配置是否正确"""
        try:
            # 创建测试对话框
            dialog = EmailTestDialog(self)
            dialog.exec_()
            
        except Exception as e:
            QMessageBox.critical(self, "错误", f"测试发送失败: {str(e)}")

    def show_help(self):
        """显示帮助对话框"""
        dialog = HelpDialog(self)
        dialog.exec_()

class EmailTestDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.email_sender = EmailSender()
        self.setup_ui()
        
    def setup_ui(self):
        self.setWindowTitle("邮箱配置测试")
        self.setMinimumWidth(400)
        
        layout = QVBoxLayout(self)
        
        # 显示当前配置信息
        config_group = QGroupBox("当前配置")
        config_layout = QFormLayout()
        
        sender_name = self.email_sender.sender_name
        sender_email = self.email_sender.sender_email
        smtp_server = self.email_sender.smtp_server
        smtp_port = str(self.email_sender.smtp_port)
        use_ssl = "是" if self.email_sender.use_ssl else "否"
        
        config_layout.addRow("发件人:", QLabel(f"{sender_name} <{sender_email}>"))
        config_layout.addRow("SMTP服务器:", QLabel(smtp_server))
        config_layout.addRow("SMTP端口:", QLabel(smtp_port))
        config_layout.addRow("使用SSL:", QLabel(use_ssl))
        
        config_group.setLayout(config_layout)
        layout.addWidget(config_group)
        
        # 测试进度和结果
        self.status_label = QLabel("准备测试...")
        layout.addWidget(self.status_label)
        
        self.progress = QProgressBar()
        self.progress.setRange(0, 3)
        self.progress.setValue(0)
        layout.addWidget(self.progress)
        
        # 按钮
        btn_layout = QHBoxLayout()
        self.test_btn = QPushButton("开始测试")
        self.test_btn.clicked.connect(self.run_test)
        self.close_btn = QPushButton("关闭")
        self.close_btn.clicked.connect(self.close)
        
        btn_layout.addWidget(self.test_btn)
        btn_layout.addWidget(self.close_btn)
        layout.addLayout(btn_layout)
        
        # 开始测试
        QTimer.singleShot(100, self.run_test)
    
    def run_test(self):
        self.test_btn.setEnabled(False)
        self.progress.setValue(0)
        
        try:
            # 测试SMTP连接
            self.status_label.setText("正在连接SMTP服务器...")
            self.progress.setValue(1)
            
            if self.email_sender.use_ssl:
                server = smtplib.SMTP_SSL(
                    self.email_sender.smtp_server, 
                    self.email_sender.smtp_port
                )
            else:
                server = smtplib.SMTP(
                    self.email_sender.smtp_server, 
                    self.email_sender.smtp_port
                )
                server.starttls()
            
            # 测试登录
            self.status_label.setText("正在验证登录信息...")
            self.progress.setValue(2)
            server.login(
                self.email_sender.sender_email, 
                self.email_sender.smtp_password
            )
            
            # 发送测试邮件
            self.status_label.setText("正在发送测试邮件...")
            self.progress.setValue(3)
            
            # 创建测试邮件
            msg = MIMEMultipart('alternative')
            msg['From'] = f"{self.email_sender.sender_name} <{self.email_sender.sender_email}>"
            msg['To'] = self.email_sender.sender_email
            msg['Subject'] = Header("邮箱配置测试", 'utf-8')
            
            html_content = f"""
            <html>
            <body>
                <h3>邮箱配置测试成功</h3>
                <p>这是一封测试邮件，用于验证邮箱配置是否正确。</p>
                <p>配置信息：</p>
                <ul>
                    <li>发件人：{self.email_sender.sender_name} &lt;{self.email_sender.sender_email}&gt;</li>
                    <li>SMTP服务器：{self.email_sender.smtp_server}</li>
                    <li>SMTP端口：{self.email_sender.smtp_port}</li>
                    <li>SSL加密：{'是' if self.email_sender.use_ssl else '否'}</li>
                </ul>
                <p>发送时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
            </body>
            </html>
            """
            
            html_part = MIMEText(html_content, 'html', 'utf-8')
            msg.attach(html_part)
            
            # 发送邮件
            server.send_message(msg)
            server.quit()
            
            # 测试完成
            self.status_label.setText("测试完成！配置正确，邮件已发送。")
            self.progress.setValue(3)
            QMessageBox.information(
                self,
                "测试成功",
                f"邮箱配置测试成功！\n已向 {self.email_sender.sender_email} 发送测试邮件。"
            )
            
        except Exception as e:
            error_msg = str(e)
            self.status_label.setText(f"测试失败: {error_msg}")
            QMessageBox.critical(
                self,
                "测试失败",
                f"邮箱配置测试失败！\n\n错误信息：{error_msg}\n\n"
                "请检查以下内容：\n"
                "1. SMTP服务器地址和端口是否正确\n"
                "2. 邮箱账号和密码是否正确\n"
                "3. 是否已开启SMTP服务\n"
                "4. 如果使用Gmail，是否已开启两步验证并使用应用专用密码"
            )
        
        finally:
            self.test_btn.setEnabled(True) 

class HelpDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("使用帮助")
        self.setMinimumSize(600, 400)
        
        layout = QVBoxLayout(self)
        
        # 创建文本浏览器
        self.help_text = QTextEdit()
        self.help_text.setReadOnly(True)
        layout.addWidget(self.help_text)
        
        # 关闭按钮
        close_btn = QPushButton("关闭")
        close_btn.clicked.connect(self.close)
        close_btn.setFixedWidth(100)
        
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        btn_layout.addWidget(close_btn)
        layout.addLayout(btn_layout)
        
        # 加载帮助文档
        self.load_help_content()
        
    def load_help_content(self):
        help_content = """
# 智能邮件群发系统使用说明

## 1. 基本使用流程

### 1.1 选择Word模板
- 点击"浏览..."选择Word文档作为邮件模板
- 在Word模板中使用 {变量名} 格式插入变量
- 变量名需要与Excel表格的列名完全一致
- 系统会自动识别模板中的所有变量

### 1.2 选择Excel数据
- 点击"浏览..."选择Excel文件
- 系统会自动识别姓名列和邮箱列
- 自动匹配Word模板中的其他变量
- 自动显示第一条数据的预览效果

### 1.3 发送邮件
1. 填写邮件主题
2. 设置发送间隔时间（秒）
3. 确认预览效果无误后点击"开始发送"
4. 可通过进度条查看发送进度
5. 如需停止发送，点击"停止发送"

## 2. 变量使用说明

### 2.1 变量格式
- 在Word中使用 {变量名} 格式
- 例如：{姓名}、{部门}、{职位}
- 变量名必须与Excel列名完全一致
- 大小写敏感，请注意保持一致

### 2.2 自动匹配规则
- 姓名列：自动匹配包含"姓名"、"名字"、"name"的列
- 邮箱列：自动匹配包含"邮箱"、"邮件"、"email"的列
- 其他变量：自动与Excel列名进行匹配
- 未匹配变量会在预览中显示 [未匹配变量: {变量名}]

## 3. 注意事项

### 3.1 文件准备
- Word模板需为.doc或.docx格式
- Excel文件需为.xls或.xlsx格式
- Excel表格第一行必须为列名
- 确保数据列名与模板变量名一致

### 3.2 发送建议
- 首次使用建议先测试邮箱配置
- 发送前请仔细检查预览效果
- 建议适当设置发送间隔时间
- 大量发送时注意邮箱服务限制

## 4. 常见问题

### 4.1 变量未匹配
- 检查变量名与Excel列名是否完全一致
- 注意大小写、空格等是否一致
- 确认Excel文件第一行是否为列名

### 4.2 邮件发送失败
- 检查邮箱配置是否正确
- 确认网络连接是否正常
- 查看是否触发发送频率限制
- 验证收件人邮箱地址是否有效

### 4.3 预览显示异常
- 确认Word模板格式是否正确
- 检查Excel数据是否完整
- 验证变量格式是否规范

## 5. 技术支持

如遇到问题，请检查：
1. 文件格式是否正确
2. 变量名是否匹配
3. 邮箱配置是否有效
4. 网络连接是否正常

如需帮助，请联系技术支持。
"""
        self.help_text.setMarkdown(help_content) 