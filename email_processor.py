import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.header import Header
import configparser
import os

class EmailSender:
    """邮件发送处理类"""
    
    def __init__(self, config_file="config.ini"):
        self.config = self._load_config(config_file)
        self.sender_name = self.config.get('EMAIL', 'sender_name')
        self.sender_email = self.config.get('EMAIL', 'sender_email')
        self.smtp_server = self.config.get('EMAIL', 'smtp_server')
        self.smtp_port = self.config.getint('EMAIL', 'smtp_port')
        self.smtp_password = self.config.get('EMAIL', 'smtp_password')
        self.use_ssl = self.config.getboolean('EMAIL', 'use_ssl')
    
    def _load_config(self, config_file):
        """加载配置文件"""
        if not os.path.exists(config_file):
            raise FileNotFoundError(f"找不到配置文件: {config_file}")
            
        config = configparser.ConfigParser()
        config.read(config_file, encoding='utf-8')
        
        # 验证必要配置
        required_options = [
            ('EMAIL', 'sender_name'),
            ('EMAIL', 'sender_email'),
            ('EMAIL', 'smtp_server'),
            ('EMAIL', 'smtp_port'),
            ('EMAIL', 'smtp_password')
        ]
        
        for section, option in required_options:
            if not config.has_option(section, option):
                raise ValueError(f"配置文件中缺少必要的选项: [{section}] {option}")
                
        return config
    
    def send_email(self, to_email, subject, html_content):
        """发送邮件"""
        # 创建邮件
        msg = MIMEMultipart('alternative')
        msg['From'] = f"{self.sender_name} <{self.sender_email}>"
        msg['To'] = to_email
        msg['Subject'] = Header(subject, 'utf-8')
        
        # 添加HTML内容
        html_part = MIMEText(html_content, 'html', 'utf-8')
        msg.attach(html_part)
        
        # 连接到SMTP服务器并发送
        try:
            if self.use_ssl:
                server = smtplib.SMTP_SSL(self.smtp_server, self.smtp_port)
            else:
                server = smtplib.SMTP(self.smtp_server, self.smtp_port)
                server.starttls()
                
            server.login(self.sender_email, self.smtp_password)
            server.send_message(msg)
            server.quit()
            
        except Exception as e:
            raise Exception(f"发送邮件失败: {str(e)}") 