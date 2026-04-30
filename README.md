# XML 表格合并工具

## 运行方式

- 开发环境直接启动：`python app.py`
- 也可以直接带文件参数启动：`python app.py left.xml right.xlsx`
- 或双击：`run_merge_tool.bat`

## 打包成程序

1. 确保本机已安装 Python 3、`PySide6`、`lxml`、`PyInstaller`
2. 在当前目录双击 `build_exe.bat`
3. 打包完成后，程序输出在：

`dist\分久必合\分久必合.exe`

## 当前打包方式

- 使用 `PyInstaller`
- 输出为 `onedir` 目录版，启动更稳，排查问题也更方便
- 入口文件：`app.py`
- 打包配置：`xml_merge_tool.spec`
- 当前图标文件：`icon.png`

## 当前支持的输入来源

- 本地 `xml`
- 本地 `xlsx`
- 本地 `csv`
- SVN XML
- Google Sheets

说明：

- `xlsx/csv` 可用于读取和比对
- 导出模板目前仍要求使用 `XML`

## 右键菜单

- 安装当前用户右键菜单：双击 `install_context_menu_current_user.bat`
- 卸载当前用户右键菜单：双击 `uninstall_context_menu_current_user.bat`

右键菜单当前会给 `.xml`、`.xlsx`、`.csv` 增加三项入口：

- `设为对比文件1`
- `与已选文件1对比`
- `清除已选文件1`

使用方式：

1. 先在第一个文件上点 `设为对比文件1`
2. 再在第二个文件上点 `与已选文件1对比`
3. 工具会自动带入两个文件并直接开始比对
