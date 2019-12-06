Tool Installation��
1��The PC that runs Windows 7 or later is required. Note that powershell.exe is the default configuration software of the Windows OS. If powershell.exe is unavailable, install one.
2��Excel 2007 or a later version has been installed on the local PC.
3��Users have enabled the Excel trust function and have permission to access object models of VBA projects. (To enable the Excel trust function, choose File > Options > Trust Center > Trust Center Settings > Macro Settings > Trust VBA project with object
model access.)
4��JAVA JDK8 or later is required. In versions later than JDK12, you need to manually generate JRE.(JAVA JDK9 or later jdk does not have jre,please excute commands under "JAVA_HOME" path in CMD "bin\jlink.exe --module-path jmods --add-modules java.base,java.logging,java.desktop --output jre")
Tool operation��
1��Run the start_EN.bat or runbat_EN.vbs file. (If the file cannot run, right-click the file and open it in the command prompt.)
2��(Optional) Click the Macro Injection tab. On the tab page, click ..., select an .xlsx file to which macros will be injected, and click Submit. Then, the tool generates an .xlsm file with the same name as the selected file in the selected directory.
3��(Optional) Click the Csv to Excel tab. On the tab page, click ..., select a .zip file to be converted, and click Submit. Then, the tool generates an .xlsm file with the same name as the selected file in the selected directory.
4��(Optional) Click the Excel to Csv tab. On the tab page, click ..., select a .xlsx or .xlsm file to be converted, and click Submit. Then, the tool generates an .zip file with the same name as the selected file in the selected directory.
5��(Optional) Click the Update Custom Package tab. On the tab page, click ..., select a .zip or .tar customization package to be updated, and click Submit. Then, the tool updates and overwrites the original customization package in the selected directory.

���߰�װ��
1. �û���ʹ��Windows7 �����ϰ汾����ϵͳ��ע�⣺powershell.exe һ��Ϊ Windows ϵͳĬ��������������û�У���Ҫ���а�װ����
2. �û��谲װjava jdk8 �����ϰ汾��jdk9 ���ϰ汾��Ҫ�ֶ�����jre, ��cmd "JAVA_HOME"Ŀ¼��ִ��������"bin\jlink.exe --module-path jmods --add-modules java.base,java.logging,java.desktop --output jre"����
3. �û�������װExcel 2007 �����ϰ汾��
4. �û��ѿ���Excel ���Σ���VBA ���̶���ģ�͵ķ���Ȩ�ޡ���������ʽ���ļ� -> ѡ�� -> �������� -> ������������ -> ������ -> ���ζ�VBA ���̶���ģ�͵ķ��ʣ���
�������У�
1��  ���С�start_CN.bat����runbat_CN.vbs���ļ��������޷����У����Ҽ���������ʾ���д򿪣�
2������ѡ��ѡ�� ����ע�롱 ҳǩ��������...����ť��ѡ������ע����Excel �ļ�����.xlsx����ʽ�����������ύ����ť�����߽�����ѡ�ļ�·���£�����һ��ͬ���Һ�׺Ϊ��.xlsm����excel �ļ���
3������ѡ��ѡ�� ��Csv תExcel ��ҳǩ��������...����ť��ѡ������ת�����ļ�����.zip����ʽ�����������ύ����ť�����߽�����ѡ�ļ�·���£�����һ��ͬ���Һ�׺Ϊ��.xlsm����excel �ļ���
4������ѡ��ѡ�� ��Excel תCsv ��ҳǩ��������...����ť��ѡ������ת�����ļ�����.xlsx����ʽ����".xlsm" ��ʽ�����������ύ����ť�����߽�����ѡ�ļ�·���£�����һ��ͬ���Һ�׺Ϊ��.zip����ѹ���ļ���
5������ѡ��ѡ�� �����¶��ư� ��ҳǩ��������...����ť��ѡ��������µĶ��ư��ļ�����.zip����.tar����ʽ�����������ύ����ť�����߽�����ѡ�ļ�·���£����²�����ԭ�еĶ��ư��ļ���