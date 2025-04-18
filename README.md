# 智能邮件群发系统

一个基于Python和PyQt5开发的智能邮件群发工具，支持Word模板和Excel数据源的自动匹配，具有现代化UI界面和友好的用户体验。

## 功能特点

- 支持Word文档作为邮件模板，完整保留格式
- 支持Excel表格作为收件人数据源
- 智能识别并自动匹配所有变量
- 自动识别姓名和邮箱列
- 实时邮件预览功能
- 自动显示变量匹配状态
- 未匹配变量智能提示
- 可配置发送时间间隔
- 发送进度实时显示
- 支持中断发送任务
- 邮箱配置测试功能
- 现代化UI界面设计

## 系统要求

- Windows 7/8/10/11 (64位)
- Microsoft Visual C++ 2015-2022 Redistributable

## 使用说明

### 1. 文件准备

#### Word模板要求
- 使用 {变量名} 格式插入变量（例如：{姓名}、{部门}）
- 变量名需要与Excel表格的列名完全一致
- 支持任意数量的变量
- 支持完整的Word格式（字体、颜色、对齐等）

示例：
```
尊敬的{姓名}：

您好！这是来自{部门}的通知。
您的职位是{职位}。

此致
敬礼
```

#### Excel数据要求
- 第一行必须是列名（将用于自动匹配变量）
- 建议包含姓名和邮箱相关列名
- 支持所有标准Excel格式

### 2. 使用流程

1. 选择Word模板
   - 点击"浏览..."选择Word文档
   - 系统自动识别模板中的所有变量
   - 自动显示变量列表

2. 选择Excel数据
   - 点击"浏览..."选择Excel文件
   - 系统自动识别姓名和邮箱列
   - 自动匹配所有变量
   - 自动显示第一条数据的预览效果

3. 发送邮件
   - 填写邮件主题
   - 设置发送间隔
   - 确认预览无误后点击"开始发送"

### 3. 变量匹配规则

#### 自动识别规则
- 姓名列：自动匹配包含"姓名"、"名字"、"name"的列
- 邮箱列：自动匹配包含"邮箱"、"邮件"、"email"的列
- 其他变量：自动与Excel列名进行匹配

#### 未匹配变量处理
- 在预览中显示为 [未匹配变量: {变量名}]
- 发送前会给出警告提示
- 可继续发送或检查修正

## 配置说明

### 邮箱配置 (config.ini)

```ini
[EMAIL]
sender_name = 发件人姓名
sender_email = your_email@example.com
smtp_server = smtp.example.com
smtp_port = 587  # 推荐使用587端口
smtp_password = your_password
use_ssl = False  # 使用587端口时设置为False，使用TLS加密
```

### 常见邮箱服务器设置

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

## 注意事项

1. 发送前检查：
   - 确保网络连接正常
   - 验证邮箱配置正确
   - 检查变量匹配状态
   - 查看预览效果

2. 发送建议：
   - 首次使用先测试邮箱配置
   - 大量发送时适当增加间隔
   - 注意邮件服务商限制
   - 定期检查发送状态

## 常见问题

1. 变量未匹配
   - 检查变量名与Excel列名是否完全一致
   - 注意大小写和空格
   - 确认Excel第一行为列名

2. 邮件格式问题
   - 确保Word文档格式正确
   - 检查变量标记格式
   - 预览中确认格式效果

3. 发送失败
   - 检查邮箱配置
   - 确认网络连接
   - 验证授权码正确性
   - 查看错误提示信息

## 版本历史

- v2.0.0
  - 新增变量自动匹配功能
  - 添加自动预览功能
  - 优化Word格式处理
  - 改进错误提示
  - 优化用户界面

## 许可说明

本项目仅供学习和参考使用。在使用本工具时，请遵守：
1. 相关法律法规
2. 邮件服务商的使用规范
3. 用户隐私保护规定
