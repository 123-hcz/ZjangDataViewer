import os
import xml.etree.ElementTree as et
from xml.dom import minidom
import json
import decimal
import pandas as pd
import openpyxl

def readConfig():
    current_dir = os.path.dirname(__file__)
    config_path = os.path.join(current_dir, "_json", "options.json")
    with open(config_path, "r", encoding="utf-8") as f:
        return json.load(f)

config = readConfig()

RW = config["usingRoundWay"]
DPN = config["decimalPlaceNum"]


def readExcel(excel_file:str,sheet:str):
    df = pd.read_excel(excel_file, sheet_name=sheet, header=None, engine='openpyxl')
    data = df.values.tolist()
    return data

def readXml(file_path: str):
    try:
        tree = et.parse(file_path)
        root = tree.getroot()
        data = []
        for row_elem in root.findall('row'):
            row_data = []
            for col_elem in row_elem.findall('col'):
                row_data.append(col_elem.text if col_elem.text is not None else "")
            data.append(row_data)
        return data
    except (et.ParseError, FileNotFoundError):
        return []

def xml_string_to_data(xml_string: str):
    """从XML字符串解析数据为列表."""
    try:
        root = et.fromstring(xml_string)
        data = []
        for row_elem in root.findall('row'):
            row_data = []
            for col_elem in row_elem.findall('col'):
                row_data.append(col_elem.text if col_elem.text is not None else "")
            data.append(row_data)
        return data
    except et.ParseError:
        return []

def getItems(Data:list,ItemRow:int):
    ItemRow -= 1
    if not Data or len(Data) <= ItemRow:
        return []
    return Data[ItemRow]

def getNames(Data:list,ItemRow:int,NameCol:int):
    ItemRow -= 1
    NameCol -= 1
    names = []
    for i in Data:
        if len(i) > NameCol:
            names.append(i[NameCol])
    if len(names) > ItemRow + 1:
        del names[:ItemRow + 1]
    return names

def getValue(Data:list,ItemRow:int,NameCol:int,ItemName:str):
    items = getItems(Data,ItemRow)
    ItemRow -= 1
    NameCol -= 1
    try:
        if ItemName is None or ItemName not in items:
            raise ValueError("项目名称未在项目中找到或是空的")
        itemIndex = items.index(ItemName)
    except ValueError as e:
        raise e
    values = []
    for i in Data:
        if len(i) > itemIndex:
            values.append(i[itemIndex])
    if len(values) > ItemRow + 1:
        del values[:ItemRow + 1]
    return values

def getNameValueDict(Names:list,Values:list):
    NameValueDic = {}
    min_len = min(len(Names), len(Values))
    for i in range(min_len):
        NameValueDic[Names[i]] = Values[i]
    return NameValueDic

def getMaxValue(NameValueDic:dict):
    Values = list(NameValueDic.values())
    float_values = []
    for v in Values:
        try:
            float_values.append(float(v))
        except (ValueError, TypeError):
            pass
    if not float_values:
        return None
    return max(float_values)

def getMaxNames(NameValueDic:dict,MaxNum:float):
    if MaxNum is None:
        return ["最大值：", "无有效数据"]
    Values = list(NameValueDic.values())
    Names = list(NameValueDic.keys())
    output = ["最大值："]
    for i in range(len(Values)):
        try:
            if float(Values[i]) == MaxNum:
                output.append(f"{Names[i]} : {Values[i]}")
        except (ValueError, TypeError):
            pass
    return output

def getMinValue(NameValueDic:dict):
    Values = list(NameValueDic.values())
    float_values = []
    for v in Values:
        try:
            float_values.append(float(v))
        except (ValueError, TypeError):
            pass
    if not float_values:
        return None
    return min(float_values)

def getMinNames(NameValueDic:dict,MinNum:float):
    if MinNum is None:
        return ["最小值：", "无有效数据"]
    Values = list(NameValueDic.values())
    Names = list(NameValueDic.keys())
    output = ["最小值："]
    for i in range(len(Values)):
        try:
            if float(Values[i]) == MinNum:
                output.append(f"{Names[i]} : {Values[i]}")
        except (ValueError, TypeError):
            pass
    return output

def getCustomizeEquation(CustomizeRule:str,ReplaceVar):
    return CustomizeRule.replace('x', ReplaceVar)

def getCustomizeValue(NameValueDic:dict,CustomizeRule:str):
    satisfyNameValues = []
    Names = list(NameValueDic.keys())
    Values = list(NameValueDic.values())
    has_compute_part = "#" in CustomizeRule
    if has_compute_part:
        parts = CustomizeRule.split("#", 1)
        rule_bool_str = parts[0].strip()
        rule_compute_str = parts[1].strip()
    else:
        rule_bool_str = CustomizeRule.strip()

    for i, val in enumerate(Values):
        if i >= len(Names): continue # 防止索引越界
        try:
            val_float = float(val)
            context = {'x': val_float}
            if eval(rule_bool_str, {}, context):
                if has_compute_part:
                    compute_result = eval(rule_compute_str, {}, context)
                    satisfyNameValues.append(f"{Names[i]} = {val} | {compute_result}")
                else:
                    satisfyNameValues.append(f"{Names[i]} = {val}")
        except (ValueError, TypeError, NameError, SyntaxError):
            continue
    return satisfyNameValues

def getAverageValue(NameValueDic:dict):
    Values = list(NameValueDic.values())
    sumA = decimal.Decimal(0)
    count = 0
    for i in Values:
        try:
            sumA += decimal.Decimal(str(i))
            count += 1
        except (decimal.InvalidOperation, TypeError, ValueError):
            pass
    if count == 0:
        return [":无有效数据"]
    average = sumA / decimal.Decimal(count)
    dpn_str = '1e-' + str(DPN)
    average_quantized = average.quantize(decimal.Decimal(dpn_str), rounding=RW)
    return [f"平均值:{average_quantized}"]

def findNames(NameValueDic:dict,Value:float):
    names = []
    for i in NameValueDic:
        try:
            if float(NameValueDic[i]) == Value:
                names.append(i)
        except (ValueError, TypeError):
            pass
    return names

def write_to_excel(data, filename):
    wb = openpyxl.Workbook()
    ws = wb.active
    for row in data:
        clean_row = [cell if cell is not None else '' for cell in row]
        ws.append(clean_row)
    wb.save(filename)

def write_to_xml(data, filename):
    root = et.Element("root")
    for row_data in data:
        row_elem = et.SubElement(root, "row")
        for cell_data in row_data:
            col_elem = et.SubElement(row_elem, "col")
            col_elem.text = str(cell_data) if cell_data is not None else ""
    xml_string = et.tostring(root, 'utf-8')
    reparsed = minidom.parseString(xml_string)
    pretty_xml = '\n'.join([line for line in reparsed.toprettyxml(indent="  ").split('\n') if line.strip()])
    with open(filename, "w", encoding="utf-8") as f:
        f.write(pretty_xml)

def data_to_xml_string(data: list):
    """将列表数据转换为XML格式的字符串."""
    root = et.Element("root")
    for row_data in data:
        row_elem = et.SubElement(root, "row")
        for cell_data in row_data:
            col_elem = et.SubElement(row_elem, "col")
            col_elem.text = str(cell_data) if cell_data is not None else ""
    return et.tostring(root, 'unicode')


if __name__ == "__main__":
    excelValue = readExcel("D:\\下载\\8年级录分(1).xlsx","Sheet")
    names = getNames(excelValue,1,3)
    values = getValue(excelValue,1,3,"语文")
    NameValueDic = getNameValueDict(names,values)

    print(getCustomizeValue(NameValueDic,getCustomizeEquation("x<=72 # x-72 # x-20 ","float(Values[i])")))