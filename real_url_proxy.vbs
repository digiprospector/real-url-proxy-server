Set WshShell = CreateObject("WScript.Shell")
Set WshEnv = WshShell.Environment("PROCESS")
currentDir = WshShell.CurrentDirectory

' 切换到脚本所在目录（确保相对路径正确）
WshShell.CurrentDirectory = currentDir

' 运行并重定向输出到日志文件
WshShell.Run "cmd /c python.exe real-url-proxy-server.py -p 5566", 0, False