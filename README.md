# excel_merger
 将多份 excel(.xlsx) 文件通过关联字段进行合并
 
# Excel合并工具使用说明

## 一、安装Python

1. **下载Python**
   - 访问Python官方网站：https://www.python.org/downloads/
   - 点击"Download Python"按钮（选择最新版本，如Python 3.9或3.10）
   - 选择适合您操作系统的安装包（Windows/Mac/Linux）

2. **安装Python**
   - Windows用户：运行下载的安装文件，勾选"Add Python to PATH"选项，然后点击"Install Now"
   - Mac用户：打开下载的.pkg文件，按照指示完成安装
   - 安装完成后，可以打开命令提示符（Windows）或终端（Mac/Linux），输入`python --version`检查是否安装成功

## 二、运行Excel合并工具

1. **准备工作**
   - 将`excel_merger.py`文件保存到您的电脑上
   - 确保您有需要合并的Excel文件

2. **启动程序**
   - Windows用户：在文件所在位置，右键点击空白处，选择"在终端中打开"，然后输入`python excel_merger.py`
   - Mac用户：打开终端，使用`cd`命令导航到文件所在目录，然后输入`python3 excel_merger.py`
   - 程序第一次运行时会自动检测并安装所需的依赖库

## 三、自动依赖安装功能

本程序具有自动检测和安装所需依赖库的功能，您无需手动安装任何额外的库。

- 首次运行时，程序会自动检查是否已安装必要的依赖库（PyQt5、pandas、openpyxl）
- 如果缺少任何依赖，会显示一个安装对话框
- 点击"安装依赖项"按钮，程序会自动下载并安装所需的库
- 安装完成后，程序会自动继续运行
- 如果您希望手动安装，可以点击"手动安装说明"查看安装命令

## 四、使用方法

1. **选择文件**
   - 点击"选择文件"按钮，选择需要合并的Excel文件（可以选择多个）

2. **设置关联关系**
   - 选择适合的关联模式：
     - 单一关联：所有表格通过一个字段关联
     - 链式关联：表格之间按照链条关联
     - 星形关联：一个中心表与其他表关联
   - 在弹出的对话框中设置具体的关联关系

3. **选择输出列**
   - 在左侧选择要包含在输出中的列
   - 可以使用搜索框快速查找列
   - 使用全选/全不选按钮快速操作

4. **输出设置**
   - 点击"选择输出路径"设置合并后文件的保存位置
   - 点击"开始合并"按钮执行合并操作
   - 合并完成后会显示成功信息

## 五、常见问题

- **如果安装依赖库失败**：确保您的电脑已连接网络，可以尝试使用"手动安装说明"中的命令手动安装
- **如果程序无法启动**：检查Python是否正确安装，可以在命令行中输入`python --version`或`python3 --version`进行确认
- **如果文件无法读取**：确保您的Excel文件格式正确且未被其他程序打开
