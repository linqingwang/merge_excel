# -*- coding: utf-8 -*-
import xlrd
import xlwt
from xlutils.copy import copy
import json
import glob

if __name__ == '__main__':
    #---获得所有等待合并文件----#
    #--合并信息---#
    with open("my.json") as f:
        json_data = json.load(f)

    sheet_NO = json_data["sheet_NO"] #文件中的第几个表格
    key_location_list = json_data["key_location"] #待合并内容的唯一标示
    value_location_list = json_data["value_location"] #子表待合并的内容
    para_location_list = json_data["para_location_list"] #子表待合并的内容
    model_file = json_data["model_file"]#合并文件的模版
    res_flie = json_data["res_file"] #合并后的文件
    src_path = json_data["src_path"]#存放待合并表格的文件夹
    allxls = glob.glob(src_path + '/*.xlsx')
    print (len(allxls))
    #--合并信息结束--#
    datavalue = []
    ori_excel_dict = {}
    content_dict = {}
    #--待保存文件
    try:
        print ("opening the file...")
        old_excel = xlrd.open_workbook(model_file)
    except Exception as e:
        print (e)
    print ("File opened.")
    res_excel = copy(old_excel)
    ws = res_excel.get_sheet(sheet_NO)
    #res_excel.add_sheet('My Sheet')
    # workbook = xlwt.Workbook()
    # worksheet = workbook.add_sheet('My Sheet')
#------创建一个字典保存key:名称与row:行数----#
    ori_sheet_num = len(old_excel.sheets())
    ori_table = old_excel.sheets()[sheet_NO]#文件的第三个表格
    for row in range(ori_table.nrows):
        ori_rowdata = ori_table.row_values(row)
        ori_key = ""
        for i in range(len(key_location_list)):
            ori_key += str(ori_rowdata[key_location_list[i]])
        ori_excel_dict[ori_key] = row
    #--reslut_contex--#

    i = 0
    while i < len(allxls):
        # print (i)
        fh = xlrd.open_workbook(allxls[i])
        sheet_num = len(fh.sheets())#表格数目
        #for shnum in range(sheet_num):#遍历每一个表格
        table = fh.sheets()[sheet_NO]#第几个表格
        for row in range(table.nrows):
            key = ""
            for i_1 in range(len(key_location_list)):
                key += str(table.row_values(row)[key_location_list[i_1]])
            rowdata = table.row_values(row)#每一行的data值
            merge_state = not bool(rowdata[value_location_list[0]])
            for j in range(len(value_location_list)):
                merge_state = merge_state and (not bool(rowdata[value_location_list[j]]))
            if not merge_state:
                for k in range(len(value_location_list)):
                    # ws.write(ori_excel_dict[key], value_location_list)
                    ws.write(ori_excel_dict[key], para_location_list[k], rowdata[value_location_list[k]])
        i += 1
        print ("表格NO.",i, "/", len(allxls))
    res_excel.save(res_flie) #合并后的文件
    print ('All done!')
