import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.header import Header
import configparser
import os
from email.utils import formataddr
from bs4 import BeautifulSoup

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
        """修改发送邮件方法以支持HTML格式"""
        try:
            # 创建邮件
            msg = MIMEMultipart('alternative')
            msg['Subject'] = Header(subject, 'utf-8')
            msg['From'] = formataddr((self.sender_name, self.sender_email))
            msg['To'] = to_email
            
            # 添加HTML内容
            html_part = MIMEText(html_content, 'html', 'utf-8')
            msg.attach(html_part)
            
            # 添加纯文本版本（从HTML中提取）
            soup = BeautifulSoup(html_content, 'html.parser')
            text_content = soup.get_text()
            text_part = MIMEText(text_content, 'plain', 'utf-8')
            msg.attach(text_part)
            
            # 连接到SMTP服务器并发送
            if self.use_ssl:
                server = smtplib.SMTP_SSL(self.smtp_server, self.smtp_port)
            else:
                server = smtplib.SMTP(self.smtp_server, self.smtp_port)
                server.ehlo()
                server.starttls()
                server.ehlo()
            
            try:
                server.login(self.sender_email, self.smtp_password)
                server.sendmail(self.sender_email, [to_email], msg.as_string())
            finally:
                server.quit()
            
        except Exception as e:
            raise Exception(f"发送邮件失败: {str(e)}")

    def send_test_email(self):
        """修改测试邮件发送逻辑"""
        try:
            # 创建邮件对象
            msg = MIMEMultipart()
            msg['Subject'] = Header('邮件发送测试', 'utf-8').encode()
            
            # 修改发件人格式处理
            sender_name = Header(self.sender_name, 'utf-8').encode()
            msg['From'] = f'{sender_name} <{self.sender_email}>'
            msg['To'] = self.sender_email
            
            # 添加正文
            msg.attach(MIMEText('这是一封测试邮件，如果您收到这封邮件，说明邮箱配置正确。', 'plain', 'utf-8'))
            
            # 连接服务器并发送
            if self.use_ssl:
                # SSL模式
                server = smtplib.SMTP_SSL(self.smtp_server, self.smtp_port)
            else:
                # TLS模式
                server = smtplib.SMTP(self.smtp_server, self.smtp_port)
                server.ehlo()  # 发送EHLO命令
                server.starttls()  # 启用TLS加密
                server.ehlo()  # TLS连接后重新发送EHLO
            
            try:
                server.login(self.sender_email, self.smtp_password)
                # 使用原始的发送方式
                server.sendmail(self.sender_email, [self.sender_email], msg.as_string())
                return True, "测试邮件发送成功！"
            finally:
                server.quit()
            
        except smtplib.SMTPAuthenticationError:
            return False, "认证失败：请检查邮箱账号和授权码是否正确"
        except smtplib.SMTPException as e:
            return False, f"SMTP错误：{str(e)}"
        except Exception as e:
            return False, f"发送失败：{str(e)}" 