# 智能邮件群发系统

一个基于Python和PyQt5开发的智能邮件群发工具，支持Word模板和Excel数据源的自动匹配，具有现代化UI界面和友好的用户体验。

## 功能特点

- 支持Word文档作为邮件模板
- 支持Excel表格作为收件人数据源
- 智能识别并自动匹配变量
- 自动识别姓名和邮箱列
- 实时邮件预览功能
- 未匹配变量智能提示
- 可配置发送时间间隔
- 发送进度实时显示
- 支持中断发送任务
- 邮箱配置测试功能
- 现代化UI界面设计
![e855ce1ce7d24461b8764f3f3875c116](https://github.com/user-attachments/assets/a8d47018-0d9c-43a7-8db5-2d9f761e45f7)

## 系统要求

- Python 3.7+
- Windows/Linux/MacOS
- Microsoft Visual C++ 14.0 或更高版本

## 快速开始

1. 克隆项目
```bash
git clone [项目地址]
cd email-sender
```

2. 创建虚拟环境（推荐）
```bash
python -m venv venv
# Windows
venv\Scripts\activate
# Linux/Mac
source venv/bin/activate
```

3. 安装依赖
```bash
pip install -r requirements.txt
```

4. 运行程序
```bash
python main.py
```

## 项目结构

```
email_sender/
│
├── main.py                # 主程序入口
├── ui.py                  # UI界面实现
├── email_processor.py     # 邮件处理模块
├── word_reader.py         # Word文档读取
├── excel_reader.py        # Excel文件读取
├── config.ini            # 配置文件
└── requirements.txt      # 依赖包列表
```

### 主要模块功能

- `main.py`: 程序入口，初始化应用
- `ui.py`: 实现图形界面和用户交互
- `email_processor.py`: 处理邮件发送逻辑
- `word_reader.py`: 处理Word模板读取
- `excel_reader.py`: 处理Excel数据读取

## 打包说明

### 环境准备
1. 安装PyInstaller
```bash
pip install pyinstaller
```

2. 确保所需资源文件存在：
- config.ini（邮箱配置文件）
- email.png（程序图标）
- README.md（说明文档）

### 打包步骤

1. 运行打包脚本
```bash
python setup.py
```

2. 打包过程说明：
- 清理旧的构建文件
- 创建版本信息
- 构建可执行文件
- 复制必要资源
- 清理临时文件

3. 打包完成后，在dist目录下可以找到：
- 邮件群发工具.exe（主程序）
- config.ini（配置文件）
- README.md（说明文档）
- email.png（程序图标）

### 打包注意事项
- 确保所有依赖包已正确安装
- 确保资源文件完整
- 需要管理员权限运行打包脚本
- 打包过程可能需要几分钟时间

## 配置说明

### 邮箱配置说明

### 1. 邮箱配置 (config.ini)

```ini
[EMAIL]
sender_name = 发件人姓名
sender_email = your_email@example.com
smtp_server = smtp.example.com
smtp_port = 587  # 推荐使用587端口
smtp_password = your_password
use_ssl = False  # 使用587端口时设置为False，使用TLS加密
```

### 2. 常见邮箱服务器设置

#### QQ邮箱
```ini
smtp_server = smtp.qq.com
smtp_port = 587
use_ssl = False
```
注意事项：
- 必须使用授权码而不是登录密码
- 在QQ邮箱设置中开启SMTP服务
- 使用生成的授权码作为smtp_password

#### 163邮箱
```ini
smtp_server = smtp.163.com
smtp_port = 587
use_ssl = False
```
注意事项：
- 必须使用授权码而不是登录密码
- 在163邮箱设置中开启SMTP服务
- 使用生成的授权码作为smtp_password

#### Gmail
```ini
smtp_server = smtp.gmail.com
smtp_port = 587
use_ssl = False
```
注意事项：
- 需要开启两步验证
- 使用应用专用密码
- 确保Google账户允许不太安全的应用访问

### 3. 常见问题解决

#### 端口相关问题
1. 465端口卡死问题：
   - 改用587端口并设置use_ssl = False
   - 确保防火墙未阻止587端口
   - 检查网络连接是否稳定

2. 587端口1000错误：
   - 检查邮箱账号和授权码是否正确
   - 确认邮箱已开启SMTP服务
   - 验证发件人地址与登录账号一致
   - 检查邮箱服务器是否支持TLS加密

#### 发送测试步骤
1. 配置文件检查：
   - 确保config.ini中的配置正确
   - 使用587端口和TLS加密
   - 验证授权码是否正确填写

2. 测试流程：
   ```
   1. 先使用邮箱客户端测试SMTP设置
   2. 确保已获取正确的授权码
   3. 使用小号测试发送
   4. 检查发送日志和错误信息
   ```

3. 故障排除：
   - 检查网络连接
   - 验证服务器地址
   - 确认账号未被限制
   - 查看详细错误日志

### 4. 安全建议

1. 发送配置：
   - 使用TLS加密（587端口）
   - 定期更新授权码
   - 避免频繁发送
   - 合理设置发送间隔

2. 账号安全：
   - 使用授权码而非密码
   - 定期检查账号活动
   - 及时处理异常登录提醒
   - 开启登录通知

### 5. 性能优化

1. 网络优化：
   - 确保网络稳定
   - 避免使用代理
   - 合理设置超时时间
   - 添加重试机制

2. 发送策略：
   - 批量发送时增加间隔
   - 避免单次发送过多
   - 监控发送速率
   - 注意服务商限制

如果仍然遇到问题，建议：
1. 查看程序日志输出
2. 记录具体错误信息
3. 尝试不同的端口配置
4. 联系技术支持获取帮助

## 使用说明

### 1. 文件准备

#### Word模板要求
- 使用 {变量名} 格式插入变量
- 变量名需要与Excel表格的列名完全一致
- 支持任意数量的变量

示例：

## 使用指南

1. 准备工作
   - 创建Word邮件模板
   - 准备Excel收件人数据
   - 配置config.ini文件

2. 启动程序
   ```bash
   python main.py
   ```

3. 操作步骤
   - 选择Word模板文件
   - 选择Excel数据文件
   - 选择姓名和邮箱列
   - 填写邮件主题
   - 设置发送间隔
   - 测试邮箱配置
   - 生成预览确认
   - 开始发送

## 注意事项

1. 发送前检查事项：
   - 确保网络连接正常
   - 验证邮箱配置正确
   - 检查模板格式无误
   - 确认收件人数据完整

2. 发送建议：
   - 首次使用建议先测试配置
   - 大量发送时适当增加间隔
   - 定期检查发送状态
   - 注意邮件服务商限制

## 常见问题解决

### 1. 打包相关
- Q: 打包失败，提示缺少依赖
  - A: 检查requirements.txt中的包是否都已安装
  - A: 尝试重新安装PyInstaller

- Q: 运行exe文件报错
  - A: 确保所有资源文件在正确位置
  - A: 检查是否缺少Visual C++运行库

### 2. 发送相关
- Q: 无法连接SMTP服务器
  - A: 检查网络连接
  - A: 验证服务器地址和端口
  - A: 确认SSL设置是否正确

- Q: 认证失败
  - A: 检查账号密码
  - A: 确认是否需要使用授权码
  - A: 验证邮箱服务是否开启SMTP

## 技术支持

如遇问题，请按以下步骤处理：
1. 检查配置文件设置
2. 查看程序运行日志
3. 确认网络连接状态
4. 提交Issue或联系技术支持

## 版本历史

- v1.0.0
  - 基础邮件发送功能
  - Word模板和Excel数据支持
  - 现代化UI界面
  - 邮箱配置测试
  - 打包功能支持

## 许可说明

本项目仅供学习和参考使用。在使用本工具时，请遵守：
1. 相关法律法规
2. 邮件服务商的使用规范
3. 用户隐私保护规定
```

