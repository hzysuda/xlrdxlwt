# coding:utf-8
import xlrd
import re
from xlrdxlwt.test import xlswrite
import sys


def getxlstable(xlsfile):
    data = xlrd.open_workbook(xlsfile)
    table = data.sheets()[0]          #通过索引顺序获取
    # table = data.sheet_by_index (0)#通过索引顺序获取
    # table = data.sheet_by_name(u'Sheet1')#通过名称获取
    nrows = table.nrows
    ncols = table.ncols
    result=[]
    for i in range(1, nrows):
        rowdata=[]
        for j in range(0, ncols):
            value=table.cell(i, j).value
            valuetype=type(value)
            if(valuetype in [int, long]):
                value=int(value)
            elif(valuetype==float):
                value=value
            elif(valuetype==bool):
                value=str(value)
            elif(valuetype.__name__=="unicode"):
                try:
                    value=value.encode("gbk")
                except Exception as ex:
                    pass
            rowdata.append(value)
        result.append(rowdata)
    return (result)

def getXlsTableByUtf8(xlsfile):
    data = xlrd.open_workbook(xlsfile,encoding_override="utf-8")
    table = data.sheets()[0]          #通过索引顺序获取
    # table = data.sheet_by_index(0) #通过索引顺序获取
    # table = data.sheet_by_name(u'Sheet1')#通过名称获取
    nrows = table.nrows
    ncols = table.ncols
    result=[]
    for i in range(1, nrows):
        rowdata=[]
        for j in range(0, ncols):

            value=table.cell(i, j).value
            valuetype=type(value)
            if(valuetype in [int, long]):
                value=int(value)
            elif(valuetype==float):
                value=value
            elif(valuetype==bool):
                value=str(value)
            elif(valuetype.__name__=="unicode"):
                try:
                    value = value.encode("utf-8")
                except Exception as ex:
                    pass

                pass
            rowdata.append(value)
        result.append(rowdata)
    return (result)

def readcsv(filepath):
    result=[]
    f=open(filepath,"r")
    fc=f.read().replace("\r","")
    f.close()
    linesz=fc.split("\n")
    if(len(linesz)>0):
        keysz=linesz[0].split(",")

        for i in range(1,len(linesz)):
            m={}
            if(linesz[i]==""):
                continue
            valuesz=linesz[i].split(",")
            for j in range(0,len(keysz)):
                m[keysz[j]]=str(valuesz[j])
            result.append(m)
    return result

def handle1(path):
    r=getxlstable(path)
    xml_data=[]

    for item in r:
        if (re.search("^[A-Za-z0-9]{6}$",str(item[1]))):
            temp_dict={}
            temp_dict["code"]=str(item[1])
            temp_dict["product_name"]=str(item[3])
            temp_dict["desc"]=str(item[4])
            temp_dict["count"]=str(item[5])
            temp_dict["size"]=str(item[6])
            temp_dict["sign1"]=str(item[7])
            temp_dict["type1"]=str(item[8])
            temp_dict["sign2"]=str(item[10])
            temp_dict["sign3"]=str(item[12])
            xml_data.append(temp_dict)
        else:continue

    print(xml_data)
    xlspath="/test/aa.xls"
    xlssavepath=sys.path[0].replace("\\","/")+xlspath
    print(xlssavepath)
    xlswrite.create_excel(xml_data, xlssavepath, fields=["code", "product_name", "desc", "count", "size", "sign1", "type1", "sign2", "sign3"])

def handle2(path):
    r=getxlstable(path)
    xml_data=[]
    for item in r:
        if (re.search("^[A-Za-z0-9]{6}$",str(item[0]))):
            temp_dict={}
            temp_dict["code"]=str(item[0])
            temp_dict["product_name"]=str(item[2])
            temp_dict["desc"]=str(item[-3])
            temp_dict["count"]=str(item[1])
            temp_dict["size"]=str(item[6])
            temp_dict["sign1"]=str(item[3])
            temp_dict["type1"]=str(item[4])
            temp_dict["sign2"]=str(item[7])
            temp_dict["sign3"]=str(item[8])
            temp_dict["unit-price"]=str(item[9])
            temp_dict["total-price"]=str(item[10])
            xml_data.append(temp_dict)
        else:continue

    print(xml_data)
    xlspath="/test/bb.xls"
    xlssavepath=sys.path[0].replace("\\","/")+xlspath
    print(xlssavepath)
    xlswrite.create_excel(xml_data, xlssavepath, fields=["code", "product_name", "desc", "count", "size", "sign1", "type1", "sign2", "sign3", "unit-price", "total-price"])

# def handle3(path):
#     r=getxlstable(path)
#     xml_data=[]
#     for item in r:
#         print(item)
#         continue
#         if (re.search("^[A-Za-z0-9]{6}$",str(item[0]))):
#             temp_dict={}
#             temp_dict["code"]=str(item[0])
#             temp_dict["product_name"]=str(item[2])
#             temp_dict["desc"]=str(item[-3])
#             temp_dict["count"]=str(item[1])
#             temp_dict["size"]=str(item[6])
#             temp_dict["sign1"]=str(item[3])
#             temp_dict["type1"]=str(item[4])
#             temp_dict["sign2"]=str(item[7])
#             temp_dict["sign3"]=str(item[8])
#             temp_dict["unit-price"]=str(item[9])
#             temp_dict["total-price"]=str(item[10])
#             xml_data.append(temp_dict)
#         else:continue
#
#     print(xml_data)
#     xlspath="/xls/cc.xls"
#     xlssavepath=sys.path[0].replace("\\","/")+xlspath
#     print(xlssavepath)
#     xlswrite.create_excel(xml_data,xlssavepath,fields=["code","product_name","desc","count","size","sign1","type1","sign2","sign3"])

if(__name__=="__main__"):
    handle1("G:/test/a.xlsx")
    handle2("G:/test/b.xls")
    # handle3("E:/xml/c.xlsx")

