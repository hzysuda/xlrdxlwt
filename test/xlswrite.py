# coding:utf-8
import random
import sys
import os
import xlrd
import xlwt
# 创建excel文件，返回创建成功与否的信息
def create_excel(data,savepath,fields=None,fields_ch=None):

    # print(fields_ch)
    result={"result":"ok"}
    # if(len(data)==0):
    #     return "fail"
    # print(len(data))
    workbook = xlwt.Workbook(encoding="utf-8")
    worksheet = workbook.add_sheet("CardList")
    style = xlwt.XFStyle()
    # 调整每行的宽度
    # for row in range(0,len(data)):
    #     worksheet.col(row).width = 6666
    # print(data)
    i = 0
    j = 0
    if(fields==None and fields_ch==None):
        keys=data[0].keys()
    else:
        keys=fields
    if fields_ch!=None:
        keys_ch=fields_ch
        for key in keys_ch:
            s = key.strip()
            # print(s)
            #print(s)
            worksheet.write(0,j, s, style)
            j += 1
    else:
        for key in keys:
            s = key.strip()
            print(s)
            #print(s)
            worksheet.write(0,j, s, style)
            j += 1
    i=1
    for rowitem in data:
        j = 0
        j=0
        for key in keys:
            if(rowitem[key]==None or rowitem[key]==""):
                s=""
            else:
                s = rowitem[key]
            #print(s)
            worksheet.write(i,j, s, style)
            j += 1
        i += 1
    workbook.save(savepath)
    return result
if(__name__=="__main__"):
    data=[]
    fields=["cardno","cardpwd"]
    fieldszh=["卡号","密码"]
    # savepath="f:/1.xls"
    savepath = "G:/test/1.xlsx"
    r=create_excel(data=data,savepath=savepath,fields=fields,fields_ch=fieldszh)
    print(r)

