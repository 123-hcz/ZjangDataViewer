
import os
import random
import xml.etree.ElementTree as et
import zipfile
import os
import json
import decimal
import pandas as pd
import openpyxl

def readConfig():
    # with open("_json/options._json","r",encoding = "utf-8") as f:
    #     return  _json.load(f)
    # 获取当前文件所在目录，并定位到 _json/options._json
    current_dir = os.path.dirname(__file__)  # 获取当前文件所在路径
    config_path = os.path.join(current_dir, "_json", "options.json")

    with open(config_path, "r", encoding="utf-8") as f:
        return json.load(f)
config = readConfig()

RW = config["usingRoundWay"]
DPN = config["decimalPlaceNum"]


def readExcel(excel_file:str,sheet:str):
    df = pd.read_excel(excel_file, sheet_name=sheet, header=None, engine='openpyxl')

    # 将 DataFrame 转换为二维列表
    data = df.values.tolist()
    return data

def getItems(Data:list,ItemRow:int):
    ItemRow -= 1
    return Data[ItemRow]

def getNames(Data:list,ItemRow:int,NameCol:int):
    ItemRow -= 1
    NameCol -= 1
    names = [ ]
    for i in Data:
        names.append(i[NameCol])
    del names[:ItemRow + 1]
    return names

def getValue(Data:list,ItemRow:int,NameCol:int,ItemName:str):
    items = getItems(Data,ItemRow)
    ItemRow -= 1
    NameCol -= 1
    try:
        itemIndex = items.index(ItemName)
    except ValueError:
        raise ValueError("ItemName not found in Item")
        return None
    values = [ ]
    for i in Data:
        values.append(i[itemIndex])
    del values[:ItemRow + 1]
    return values

def getNameValueDict(Names:list,Values:list):
    NameValueDic = { }
    for i in range(len(Values)):
        NameValueDic[Names[i]] = Values[i]
    return NameValueDic

def getMaxValue(NameValueDic:dict):
    Values = list(NameValueDic.values())
    for i in range(len(Values)):
        try:
            Values[i] = float(Values[i])
        except TypeError:
            raise TypeError(f"{Values[i]} is not a float")
        except ValueError:
            raise ValueError(f"{Values[i]} is not a float")
    return max(Values)

def getMaxNames(NameValueDic:dict,MaxNum:float):
    Values = list(NameValueDic.values())
    Names = list(NameValueDic.keys())
    output = ["最大值："]
    for i in range(len(Values)):
        if float(Values[i]) == MaxNum:
            output.append(Names[i]+ " : " + str(Values[i]))
    return output


def getMinValue(NameValueDic:dict):
    Values = list(NameValueDic.values())
    for i in range(len(Values)):
        try:
            Values[i] = float(Values[i])
        except TypeError:
            raise TypeError(f"{Values[i]} is not a float")
        except ValueError:
            raise ValueError(f"{Values[i]} is not a float")
    return min(Values)

def getMinNames(NameValueDic:dict,MinNum:float):
    Values = list(NameValueDic.values())
    Names = list(NameValueDic.keys())
    output = ["最小值："]
    for i in range(len(Values)):
        if float(Values[i]) == MinNum:
            output.append(Names[i]+ " : " + str(Values[i]))
    return output

def getCustomizeEquation(CustomizeRule:str,ReplaceVar):
    CustomizeRule = list(CustomizeRule)
    for i in range(len(CustomizeRule)):
        if CustomizeRule[i] == "x":
            CustomizeRule[i] = ReplaceVar
    return  "".join(CustomizeRule)

def getCustomizeCompute(NameValueDic:dict,CustomizeRule:str):
    Computes = CustomizeRule.split("#")[1:]
    for i in range(len(Computes)):
        Computes[i] = eval(Computes[i])
    return Computes


def getCustomizeValue(NameValueDic:dict,CustomizeRule:str):
    if "#" in CustomizeRule:
        CustomizeRuleToFloat = CustomizeRule.split("#")[1]
        CustomizeRuleToBool = CustomizeRule.split("#")[0]
    else:
        CustomizeRuleToBool = CustomizeRule
    satisfyNameValues = [ ]
    Names = list(NameValueDic.keys())
    Values = list(NameValueDic.values())
    if "#" in CustomizeRule:
        for i in range(len(Values)):
            if eval(CustomizeRuleToBool):
                satisfyNameValues.append(f"{Names[i]} = {Values[i]} |{eval(CustomizeRuleToFloat)}")
    else:
        for i in range(len(Values)):
            if eval(CustomizeRuleToBool):
                satisfyNameValues.append(f"{Names[i]} = {Values[i]}")
    return satisfyNameValues

def getAverageValue(NameValueDic:dict):
    '''
    ROUND_UP: 远离零方向舍入，即总是增加绝对值。
    //向上取整

    示例：Decimal('1.3').quantize(Decimal('1'), rounding=ROUND_UP) 结果为 2
    示例：Decimal('-1.3').quantize(Decimal('1'), rounding=ROUND_UP) 结果为 -2

    ROUND_DOWN: 靠近零方向舍入，即总是减少绝对值。
    //向下取整

        示例：Decimal('1.7').quantize(Decimal('1'), rounding=ROUND_DOWN) 结果为 1
        示例：Decimal('-1.7').quantize(Decimal('1'), rounding=ROUND_DOWN) 结果为 -1

    ROUND_CEILING: 向正无穷方向舍入，只对负数有效。[不必要]
    //负数向上取整

        示例：Decimal('-1.1').quantize(Decimal('1'), rounding=ROUND_CEILING) 结果为 -1

    ROUND_FLOOR: 向负无穷方向舍入，只对正数有效。
    //正数向下取整

        示例：Decimal('1.1').quantize(Decimal('1'), rounding=ROUND_FLOOR) 结果为 1

    ROUND_HALF_UP: 如果舍弃部分 >= 0.5，则向上舍入；否则向下舍入。
    //四舍五入

        示例：Decimal('1.5').quantize(Decimal('1'), rounding=ROUND_HALF_UP) 结果为 2
        示例：Decimal('2.4').quantize(Decimal('1'), rounding=ROUND_HALF_UP) 结果为 2

    ROUND_HALF_DOWN: 如果舍弃部分 > 0.5，则向上舍入；否则向下舍入。
    //五舍六入

        示例：Decimal('1.5').quantize(Decimal('1'), rounding=ROUND_HALF_DOWN) 结果为 2
        示例：Decimal('2.5').quantize(Decimal('1'), rounding=ROUND_HALF_DOWN) 结果为 2

    ROUND_HALF_EVEN: 银行家舍入法，如果舍弃部分 = 0.5，则选择最接近的偶数。
    //银行家舍入

    示例：Decimal('2.5').quantize(Decimal('1'), rounding=ROUND_HALF_EVEN) 结果为 2
    示例：Decimal('3.5').quantize(Decimal('1'), rounding=ROUND_HALF_EVEN) 结果为 4
    '''
    Values = list(NameValueDic.values())
    sumA = 0
    for i in Values:
        try:
            addValue = decimal.Decimal(str(i))
            sumA = decimal.Decimal(sumA + addValue)
        except:
            raise TypeError(f"{i} is not a float/int")
    sumA = decimal.Decimal(str(sumA))
    count = decimal.Decimal(str(len(Values)))
    average = decimal.Decimal(sumA / count)
    dpn = str(1 / (10 ** int(DPN)))
    outputMessage =  [f":{average.quantize(decimal.Decimal(dpn), rounding=RW)}"]
    return outputMessage

def findNames(NameValueDic:dict,Value:float):
    names = [ ]

    for i in NameValueDic:
        try:
            if float(NameValueDic[i]) == Value:
                names.append(i)
        except TypeError:
            raise TypeError(f"{NameValueDic[i]} is not a float")
        except ValueError:
            raise ValueError(f"{Value[i]} is not a float")
    return names

def write_to_excel(data, filename):
    wb = openpyxl.Workbook()
    ws = wb.active
    for row in data:
        ws.append(row)
    wb.save(filename)


if __name__ == "__main__":
    excelValue = readExcel("D:\下载\8年级录分(1).xlsx","Sheet")
    names = getNames(excelValue,1,3)
    values = getValue(excelValue,1,3,"语文")
    NameValueDic = getNameValueDict(names,values)

    print(getCustomizeValue(NameValueDic,getCustomizeEquation("x<=72 # x-72 # x-20 ","float(Values[i])")))
    print(getCustomizeCompute(NameValueDic, getCustomizeEquation("x<=72 # x-72 # x-20 ", "float(Values[i])")))