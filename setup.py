import PyInstaller.__main__
import os
import shutil
import sys
from datetime import datetime
import site
import PyQt5

def get_pyqt_path():
    """获取PyQt5安装路径"""
    return os.path.dirname(PyQt5.__file__)

def clean_dist():
    """清理dist目录"""
    if os.path.exists('dist'):
        shutil.rmtree('dist')
    os.makedirs('dist')

def clean_build():
    """清理build目录"""
    if os.path.exists('build'):
        shutil.rmtree('build')
    if os.path.exists('*.spec'):
        try:
            os.remove('*.spec')
        except:
            pass

def copy_resources():
    """复制必要的资源文件"""
    resource_files = [
        'config.ini',
        'README.md',
        'email.png'  # 程序图标
    ]
    
    for file in resource_files:
        if os.path.exists(file):
            shutil.copy(file, 'dist/')
        else:
            print(f"警告: {file} 文件不存在")

def create_version_info():
    """创建版本信息文件"""
    version_info = f"""
VSVersionInfo(
    ffi=FixedFileInfo(
        filevers=(2, 0, 0, 0),
        prodvers=(2, 0, 0, 0),
        mask=0x3f,
        flags=0x0,
        OS=0x40004,
        fileType=0x1,
        subtype=0x0,
        date=(0, 0)
    ),
    kids=[
        StringFileInfo([
            StringTable(
                u'080404b0',
                [
                    StringStruct(u'CompanyName', u'Your Company'),
                    StringStruct(u'FileDescription', u'智能邮件群发系统'),
                    StringStruct(u'FileVersion', u'2.0.0'),
                    StringStruct(u'InternalName', u'email_sender'),
                    StringStruct(u'LegalCopyright', u'Copyright (C) {datetime.now().year}'),
                    StringStruct(u'OriginalFilename', u'邮件群发工具.exe'),
                    StringStruct(u'ProductName', u'智能邮件群发系统'),
                    StringStruct(u'ProductVersion', u'2.0.0')
                ])
        ]),
        VarFileInfo([VarStruct(u'Translation', [2052, 1200])])
    ]
)
"""
    with open('version_info.txt', 'w', encoding='utf-8') as f:
        f.write(version_info)

def build_executable():
    """构建可执行文件"""
    pyqt_path = get_pyqt_path()
    
    # 构建命令列表
    command = [
        'main.py',                        # 主脚本
        '--name=邮件群发工具',            # 程序名称
        '--windowed',                     # 使用窗口模式
        '--onefile',                      # 打包成单个文件
        '--icon=email.png',              # 程序图标
        '--version-file=version_info.txt', # 版本信息
        '--add-data=config.ini;.',        # 配置文件
        '--add-data=README.md;.',         # 说明文档
        '--add-data=email.png;.',         # 图标文件
        '--clean',                        # 清理临时文件
        '--noconfirm',                    # 不询问确认
        '--uac-admin',                    # 请求管理员权限
        '--noupx',                        # 不使用UPX压缩
        f'--workpath=build',              # 指定构建目录
        f'--distpath=dist',               # 指定输出目录
        
        # 添加必要的隐式导入
        '--hidden-import=PyQt5.sip',
        '--hidden-import=PyQt5.QtCore',
        '--hidden-import=PyQt5.QtGui',
        '--hidden-import=PyQt5.QtWidgets',
        '--hidden-import=docx',
        '--hidden-import=docx.opc.exceptions',
        '--hidden-import=docx.shared',
        '--hidden-import=bs4',
        '--hidden-import=lxml._elementpath',
        '--hidden-import=lxml.etree',
        
        # 收集所有必要的包
        '--collect-all=docx',
        '--collect-all=lxml',
        '--collect-all=bs4',
        
        # 排除不需要的Qt模块
        '--exclude-module=PyQt6',
        '--exclude-module=PySide6',
        '--exclude-module=PySide2',
    ]
    
    # 添加PyQt5依赖
    qt_path = os.path.join(os.path.dirname(PyQt5.__file__), 'Qt5')
    if os.path.exists(qt_path):
        # 添加Qt5的bin目录
        bin_path = os.path.join(qt_path, 'bin')
        if os.path.exists(bin_path):
            command.append(f'--add-data={bin_path};PyQt5/Qt5/bin')
        
        # 添加Qt5的plugins目录
        plugins_path = os.path.join(qt_path, 'plugins')
        if os.path.exists(plugins_path):
            command.append(f'--add-data={plugins_path};PyQt5/Qt5/plugins')
    
    # 添加python-docx依赖
    try:
        import docx
        docx_path = os.path.dirname(docx.__file__)
        command.append(f'--add-data={docx_path};docx')
    except ImportError:
        print("警告: python-docx模块未找到")
        
    # 添加beautifulsoup4依赖
    try:
        import bs4
        bs4_path = os.path.dirname(bs4.__file__)
        command.append(f'--add-data={bs4_path};bs4')
    except ImportError:
        print("警告: beautifulsoup4模块未找到")

    # 运行构建命令
    PyInstaller.__main__.run(command)

def main():
    """主函数"""
    try:
        print("开始构建应用...")
        print("1. 清理旧文件...")
        clean_dist()
        clean_build()
        
        print("2. 创建版本信息...")
        create_version_info()
        
        print("3. 构建可执行文件...")
        os.environ['PYTHONPATH'] = os.path.dirname(os.path.abspath(__file__))
        build_executable()
        
        print("4. 复制资源文件...")
        copy_resources()
        
        print("5. 清理临时文件...")
        if os.path.exists('version_info.txt'):
            os.remove('version_info.txt')
        
        print("构建完成！输出目录: dist/")
        
    except Exception as e:
        print(f"构建失败: {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main() 