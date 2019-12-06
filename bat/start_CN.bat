@echo off

if exist "%JAVA_HOME%" (
@echo "%JAVA_HOME%"
"%JAVA_HOME%"\jre\bin\java.exe -Xmx512m -Xms128m -jar excel-adapter.jar CN normal 50
) else (
@echo "please install jdk1.8 or updated and config env JAVA_HOME"
pause
)