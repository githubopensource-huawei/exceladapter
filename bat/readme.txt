Tool Installation：
1、The PC that runs Windows 7 or later is required. Note that powershell.exe is the default configuration software of the Windows OS. If powershell.exe is unavailable, install one.
2、Excel 2007 or a later version has been installed on the local PC.
3、Users have enabled the Excel trust function and have permission to access object models of VBA projects. (To enable the Excel trust function, choose File > Options > Trust Center > Trust Center Settings > Macro Settings > Trust VBA project with object
model access.)
4、JAVA JDK8 or later is required. In versions later than JDK12, you need to manually generate JRE.(JAVA JDK9 or later jdk does not have jre,please excute commands under "JAVA_HOME" path in CMD "bin\jlink.exe --module-path jmods --add-modules java.base,java.logging,java.desktop --output jre")
Tool operation：
1、Run the start_EN.bat or runbat_EN.vbs file. (If the file cannot run, right-click the file and open it in the command prompt.)
2、(Optional) Click the Macro Injection tab. On the tab page, click ..., select an .xlsx file to which macros will be injected, and click Submit. Then, the tool generates an .xlsm file with the same name as the selected file in the selected directory.
3、(Optional) Click the Csv to Excel tab. On the tab page, click ..., select a .zip file to be converted, and click Submit. Then, the tool generates an .xlsm file with the same name as the selected file in the selected directory.
4、(Optional) Click the Excel to Csv tab. On the tab page, click ..., select a .xlsx or .xlsm file to be converted, and click Submit. Then, the tool generates an .zip file with the same name as the selected file in the selected directory.
5、(Optional) Click the Update Custom Package tab. On the tab page, click ..., select a .zip or .tar customization package to be updated, and click Submit. Then, the tool updates and overwrites the original customization package in the selected directory.

工具安装：
1. 用户需使用Windows7 或以上版本操作系统（注意：powershell.exe 一般为 Windows 系统默认配置软件，如果没有，需要自行安装）。
2. 用户需安装java jdk8 或以上版本（jdk9 以上版本需要手动生成jre, 在cmd "JAVA_HOME"目录下执行命令行"bin\jlink.exe --module-path jmods --add-modules java.base,java.logging,java.desktop --output jre"）。
3. 用户本机安装Excel 2007 或以上版本。
4. 用户已开启Excel 信任，对VBA 工程对象模型的访问权限。（开启方式：文件 -> 选项 -> 信任中心 -> 信任中心设置 -> 宏设置 -> 信任对VBA 工程对象模型的访问）。
工具运行：
1、  运行“start_CN.bat”或“runbat_CN.vbs”文件。（如无法运行，可右键在命令提示符中打开）
2、（可选）选择 “宏注入” 页签，单击“...”按钮，选择所需注入宏的Excel 文件（“.xlsx”格式），单击“提交”按钮，工具将在所选文件路径下，生成一个同名且后缀为“.xlsm”的excel 文件。
3、（可选）选择 “Csv 转Excel ”页签，单击“...”按钮，选择所需转换的文件（“.zip”格式），单击“提交”按钮，工具将在所选文件路径下，生成一个同名且后缀为“.xlsm”的excel 文件。
4、（可选）选择 “Excel 转Csv ”页签，单击“...”按钮，选择所需转换的文件（“.xlsx”格式或者".xlsm" 格式），单击“提交”按钮，工具将在所选文件路径下，生成一个同名且后缀为“.zip”的压缩文件。
5、（可选）选择 “更新定制包 ”页签，单击“...”按钮，选择所需更新的定制包文件（“.zip”或“.tar”格式），单击“提交”按钮，工具将在所选文件路径下，更新并覆盖原有的定制包文件。