import excel
import os

config = excel.readConfig()

filePath = None

def runCommand(command):
    global filePath
    command = command.split(" ")
    if command[0] == "exit":
        exit()
    elif command[0] == "help":
        print("""
        exit: 退出程序
        help: 显示帮助信息
        open <filePath>: 打开文件
        """)
    elif command[0] == "cd":
        try:
            os.chdir(command[1])
            filePath = command[1]
        except:
            print("[错误] 无效的目录")
    elif command[0] == "open":
        try:
            f = excel.readExcel(command[1],command[2])
            print(f)
        except:
            print("[错误] 无效的文件")




while True:
    command = input(f"ZDV.2.2.3 & {filePath} > ")
    runCommand(command)
