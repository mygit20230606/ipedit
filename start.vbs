Set shell = CreateObject("Shell.Application")
'若要显示cmd 就把0改为1
shell.ShellExecute "cmd.exe", "/c cd /d E:\wj\project\ipedit && .venv\Scripts\activate && python main.py", "", "runas", 0