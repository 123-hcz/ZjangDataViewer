import excel
import os

config = excel.readConfig()

filePath = None

class Log():
    LogError = 1
    LogInfo = 2
    LogDebug = 3
    LogWarning = 4

    def __init__():
        pass

    
    def log(information,level):
        if level == 1:
            print(f"[错误] {information}\n")
        elif level == 2:
            print(f"[信息] {information}\n")
        elif level == 3:
            print(f"[调试] {information}\n")
        elif level == 4:
            print(f"[警告] {information}\n")

L = Log


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
            L.log(f"目录 {command[1]} 不存在",L.LogError)

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
            
            L.log(f"打开文件 {command[1]} 成功",L.LogInfo)

        except:
            L.log("打开文件失败",L.LogError)
    
    elif command[0] == "get":
        if command[1] == "max":
            pass

        elif command[1] == "min":
            pass

        elif command[1] == "average":
            pass

        elif command[1] == "customize":
            pass
        
        elif command[1] == "datavaluedict":
            pass





while True:
    command = input(f"ZDV.2.2.3 & {filePath} > ")
    runCommand(command)

