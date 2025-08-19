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
        if os.path.exists(command[1]):
            filePath = command[1]
        else:
            print("[错误] 无效的目录")

    elif command[0] == "open":
        try:
            if filePath is None:
                f = excel.readExcel(command[1],command[2])


            elif filePath.split(".")[-1] == "xlsx":
                f = excel.readExcel(filePath,command[1])

            for i in f:
                for j in i:
                    print(j, end="|\t")
                print("\n")

        except:
            print("[错误] 无效的文件")





while True:
    command = input(f"ZDV.2.2.3 & {filePath} > ")
    runCommand(command)
