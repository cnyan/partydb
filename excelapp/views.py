# -*- coding: utf-8 -*-
# 上一句话是要识别中文
from __future__ import unicode_literals

import json
from io import BytesIO
import numpy as np
from django.http import JsonResponse, StreamingHttpResponse
from django.shortcuts import HttpResponse
import xlrd  # 读取excel
import xlwt  # 写入excel
import datetime
from .models import excelUser
from django.db import connection
import os


# Create your views here.


def index(request):
    res = 'excel app '
    return HttpResponse(res)


def all_path(dirname):
    result = []
    for maindir, subdir, file_name_list in os.walk(dirname):
        for filename in file_name_list:
            apath = os.path.join(maindir, filename)
            result.append(apath)
    # print(result)
    # 列表排序
    result.sort(key=lambda x: int(x.split('\\')[-1].split('.')[0]))
    return result  # 返回文件列表


# 读取党务表1文件夹
def read_work1(request):
    rootdir = 'D:/home/闫继龙/党务/2018年9月党统/计算汇总/表1'
    excel_list = all_path(rootdir)
    row_dict = {}

    # 所有数据
    work_data = np.zeros((3, 12), dtype='int16')

    # excel_list = ['D:/home/闫继龙/党务/2018年9月党统/计算汇总/表1/1.xls']
    for file_excel in excel_list:
        by_sheet = u'附件3—党组织情况统计表（表1）'
        # 打开数据所在的工作簿，以及选择存有数据的工作表
        book = xlrd.open_workbook(file_excel)
        sheet = book.sheet_by_name(by_sheet)
        n_rows = sheet.nrows  # 行数

        sheet_data = []  # 表数据
        # data_sheet = np.zeros((3, 12))
        # 按行遍历一张表
        for row_num in range(6, 9):
            row = sheet.row_values(row_num)
            row_dict[row_num] = row

            row_data = []  # 行数据
            col_list = [4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15]
            for c in range(len(col_list)):
                df = row[col_list[c]]
                if isinstance(df, str):
                    df = 0
                else:
                    df = int(df)
                row_data.append(df)
            sheet_data.append(row_data)

        sheet_data = np.array(sheet_data)
        # print(sheet_data)
        work_data += sheet_data
        # print('88888888888888' * 8)
    print(work_data)

    da = {
        'code': '200',
        'msg': '成功',
        'data': work_data.tolist()
    }
    return HttpResponse(json.dumps(da, ensure_ascii=False), content_type="application/json,charset=utf-8")


# 读取党务表1文件夹
def read_work2(request):
    rootdir = 'D:/home/闫继龙/党务/2018年9月党统/计算汇总/表2'

    excel_list = all_path(rootdir)
    row_dict = {}

    # 所有数据
    work_data = np.zeros((10, 19), dtype='int16')

    # excel_list = ['D:/home/闫继龙/党务/2018年9月党统/计算汇总/表1/1.xls']
    for file_excel in excel_list:
        by_sheet = u'附件4—教师党员结构统计表（表2）'
        # 打开数据所在的工作簿，以及选择存有数据的工作表
        book = xlrd.open_workbook(file_excel)
        sheet = book.sheet_by_name(by_sheet)
        n_rows = sheet.nrows  # 行数

        sheet_data = []  # 表数据
        # data_sheet = np.zeros((3, 12))
        # 按行遍历一张表
        for row_num in range(7, 17):
            row = sheet.row_values(row_num)
            row_dict[row_num] = row

            row_data = []  # 行数据
            col_list = [3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21]
            for c in range(len(col_list)):
                df = row[col_list[c]]

                if isinstance(df, str):
                    df = 0
                else:
                    df = int(df)
                row_data.append(df)
            sheet_data.append(row_data)

        sheet_data = np.array(sheet_data)
        # print(sheet_data)
        work_data += sheet_data
        # print('88888888888888' * 8)
    print(work_data)

    da = {
        'code': '200',
        'msg': '成功',
        'data': work_data.tolist()
    }
    return HttpResponse(json.dumps(da, ensure_ascii=False), content_type="application/json,charset=utf-8")


# 读取党务表1文件夹
def read_work3(request):
    rootdir = 'D:/home/闫继龙/党务/2018年9月党统/计算汇总/表3'

    excel_list = all_path(rootdir)
    row_dict = {}

    # 所有数据
    work_data = np.zeros((4, 10), dtype='int16')

    # excel_list = ['D:/home/闫继龙/党务/2018年9月党统/计算汇总/表3/1.xls']
    for file_excel in excel_list:
        by_sheet = u'附件5—高层次人才党员统计表（表3）'
        # 打开数据所在的工作簿，以及选择存有数据的工作表
        book = xlrd.open_workbook(file_excel)
        sheet = book.sheet_by_name(by_sheet)
        n_rows = sheet.nrows  # 行数

        sheet_data = []  # 表数据

        # 按行遍历一张表
        for row_num in range(6, 10):
            row = sheet.row_values(row_num)
            row_dict[row_num] = row
            # print(row[3:13])
            row_data = []  # 行数据
            col_list = [3, 4, 5, 6, 7, 8, 9, 10, 11, 12]
            for c in range(len(col_list)):
                df = row[col_list[c]]
                if isinstance(df, str):
                    df = 0
                else:
                    df = int(df)
                row_data.append(df)
            sheet_data.append(row_data)

        sheet_data = np.array(sheet_data)
        # print(sheet_data)
        work_data = work_data + sheet_data

        # print('88888888888888' * 8)
    print(work_data)

    da = {
        'code': '200',
        'msg': '成功',
        'data': work_data.tolist()
        # 'data': row_dict
    }
    return HttpResponse(json.dumps(da, ensure_ascii=False), content_type="application/json,charset=utf-8")


# 读取党务表1文件夹
def read_work4(request):
    rootdir = 'D:/home/闫继龙/党务/2018年9月党统/计算汇总/表4'

    excel_list = all_path(rootdir)
    row_dict = {}

    # 所有数据
    work_data = np.zeros((3, 9), dtype='int16')

    # excel_list = ['D:/home/闫继龙/党务/2018年9月党统/计算汇总/表3/1.xls']
    for file_excel in excel_list:
        by_sheet = u'附件6—“双带头人”党支部书记配备情况统计表（表4）'
        # 打开数据所在的工作簿，以及选择存有数据的工作表
        book = xlrd.open_workbook(file_excel)
        sheet = book.sheet_by_name(by_sheet)
        n_rows = sheet.nrows  # 行数

        sheet_data = []  # 表数据

        # 按行遍历一张表
        for row_num in range(5, 8):
            row = sheet.row_values(row_num)
            row_dict[row_num] = row
            # print(row[3:13])
            row_data = []  # 行数据
            col_list = [3, 4, 5, 6, 7, 8, 9, 10, 11]
            for c in range(len(col_list)):
                df = row[col_list[c]]
                if isinstance(df, str):
                    df = 0
                else:
                    df = int(df)
                row_data.append(df)
            sheet_data.append(row_data)

        sheet_data = np.array(sheet_data)
        # print(sheet_data)
        work_data = work_data + sheet_data

        # print('88888888888888' * 8)
    print(work_data)

    da = {
        'code': '200',
        'msg': '成功',
        'data': work_data.tolist()
        # 'data': row_dict
    }
    return HttpResponse(json.dumps(da, ensure_ascii=False), content_type="application/json,charset=utf-8")


# 读取党务表1文件夹
def read_work5(request):
    rootdir = 'D:/home/闫继龙/党务/2018年9月党统/计算汇总/表5'

    excel_list = all_path(rootdir)
    row_dict = {}

    # 所有数据
    work_data = np.zeros((16, 10), dtype='int16')

    # excel_list = ['D:/home/闫继龙/党务/2018年9月党统/计算汇总/表1/1.xls']
    for file_excel in excel_list:
        by_sheet = u'附件7—学生党员结构和党组织统计表（表5）'
        # 打开数据所在的工作簿，以及选择存有数据的工作表
        book = xlrd.open_workbook(file_excel)
        sheet = book.sheet_by_name(by_sheet)
        n_rows = sheet.nrows  # 行数

        sheet_data = []  # 表数据
        # data_sheet = np.zeros((3, 12))
        # 按行遍历一张表
        for row_num in range(7, 23):
            row = sheet.row_values(row_num)
            row_dict[row_num] = row

            row_data = []  # 行数据
            col_list = [ 4, 5, 6, 7, 8, 9, 10, 11, 12, 13]
            for c in range(len(col_list)):
                df = row[col_list[c]]
                if isinstance(df, str):
                    df = 0
                else:
                    df = int(df)
                row_data.append(df)
            sheet_data.append(row_data)

        sheet_data = np.array(sheet_data)
        # print(sheet_data)
        work_data += sheet_data
        # print('88888888888888' * 8)
    print(work_data)

    da = {
        'code': '200',
        'msg': '成功',
        'data': work_data.tolist()
    }
    return HttpResponse(json.dumps(da, ensure_ascii=False), content_type="application/json,charset=utf-8")


# 读取党务表1文件夹
def read_work6(request):
    rootdir = 'D:/home/闫继龙/党务/2018年9月党统/计算汇总/表6'

    excel_list = all_path(rootdir)
    row_dict = {}

    # 所有数据
    work_data = np.zeros((4, 11), dtype='int16')

    # excel_list = ['D:/home/闫继龙/党务/2018年9月党统/计算汇总/表3/1.xls']
    for file_excel in excel_list:
        by_sheet = u'附件8—失联党员情况汇总表（表6）'
        # 打开数据所在的工作簿，以及选择存有数据的工作表
        book = xlrd.open_workbook(file_excel)
        sheet = book.sheet_by_name(by_sheet)
        n_rows = sheet.nrows  # 行数

        sheet_data = []  # 表数据

        # 按行遍历一张表
        for row_num in range(6, 10):
            row = sheet.row_values(row_num)
            row_dict[row_num] = row
            # print(row[3:13])
            row_data = []  # 行数据
            col_list = [3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13]
            for c in range(len(col_list)):
                df = row[col_list[c]]
                if isinstance(df, str):
                    df = 0
                else:
                    df = int(df)
                row_data.append(df)
            sheet_data.append(row_data)

        sheet_data = np.array(sheet_data)
        # print(sheet_data)
        work_data = work_data + sheet_data

        # print('88888888888888' * 8)
    print(work_data)

    da = {
        'code': '200',
        'msg': '成功',
        'data': work_data.tolist()
        # 'data': row_dict
    }
    return HttpResponse(json.dumps(da, ensure_ascii=False), content_type="application/json,charset=utf-8")


# 读取党务表1文件夹
def read_work7(request):
    rootdir = 'D:/home/闫继龙/党务/2018年9月党统/计算汇总/表7'

    excel_list = all_path(rootdir)
    row_dict = {}

    # 所有数据
    work_data = np.zeros((4, 8), dtype='int16')

    # excel_list = ['D:/home/闫继龙/党务/2018年9月党统/计算汇总/表3/1.xls']
    for file_excel in excel_list:
        by_sheet = u'附件9-失联党员组织处置情况汇总表（表7）'
        # 打开数据所在的工作簿，以及选择存有数据的工作表
        book = xlrd.open_workbook(file_excel)
        sheet = book.sheet_by_name(by_sheet)
        n_rows = sheet.nrows  # 行数

        sheet_data = []  # 表数据

        # 按行遍历一张表
        for row_num in range(6, 10):
            row = sheet.row_values(row_num)
            row_dict[row_num] = row
            # print(row[3:13])
            row_data = []  # 行数据
            col_list = [2, 3, 4, 5, 6, 7, 8, 9]
            for c in range(len(col_list)):
                df = row[col_list[c]]
                if isinstance(df, str):
                    df = 0
                else:
                    df = int(df)
                row_data.append(df)
            sheet_data.append(row_data)

        sheet_data = np.array(sheet_data)
        # print(sheet_data)
        work_data = work_data + sheet_data

        # print('88888888888888' * 8)
    print(work_data)

    da = {
        'code': '200',
        'msg': '成功',
        'data': work_data.tolist()
        # 'data': row_dict
    }
    return HttpResponse(json.dumps(da, ensure_ascii=False), content_type="application/json,charset=utf-8")


# 读取excel，并写入数据库
def write_excel(request):
    file_excel = 'D:/home/班级数据/data2/1.xls'
    by_sheet = u'附件3—党组织情况统计表（表1）'
    # 打开数据所在的工作簿，以及选择存有数据的工作表
    book = xlrd.open_workbook(file_excel)
    sheet = book.sheet_by_name(by_sheet)
    n_rows = sheet.nrows  # 行数
    row_dict = {}

    for row_num in range(1, n_rows):
        row = sheet.row_values((row_num))

        seq_row = {'seq': row[0], 'name': row[1], 'major': row[2], 'idcard': row[3]}
        # 向数据库保存
        # models.excelUser.objects.create(seq=row[0], name=row[1], major=row[2], idcard=row[3])
        row_dict[row_num] = row

    da = {
        'code': '200',
        'msg': '成功',
        'data': row_dict
    }
    return HttpResponse(json.dumps(da, ensure_ascii=False), content_type="application/json,charset=utf-8")


# 读取数据库，并写入excel
def read_excel(request):
    # 创建工作簿
    wb = xlwt.Workbook(encoding='utf-8')
    # 创建表
    ws = wb.add_sheet('Sheet1')
    row_num = 0
    font_style = xlwt.XFStyle()
    # 二进制
    font_style.font.bold = True
    # 表头内容
    columns = ['序号', '姓名', '专业', 'idcard']
    row_dict = {}

    # 写进表头内容
    for col_num in range(len(columns)):
        ws.write(row_num, col_num, columns[col_num], font_style)

    # Sheet body, remaining rows
    font_style = xlwt.XFStyle()
    # 获取数据库数据
    rows = excelUser.objects.values_list('seq', 'name', 'major', 'idcard')

    # 遍历提取出来的内容
    for row in rows:
        row_dict[row] = row
        row_num += 1
        # 逐行写入Excel
        for col_num in range(len(row)):
            ws.write(row_num, col_num, row[col_num], font_style)

    # 实现下载
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    response = StreamingHttpResponse(output)
    response['content_type'] = 'application/vnd.ms-excel'
    response['charset'] = 'utf-8'
    response['Content-Disposition'] = 'attachment; filename="{0}.xls"'.format(
        str(datetime.datetime.now().strftime('%Y%m%d%H%M')))
    return response


# 删除表中所有数据
def delete_table(request):
    cursor = connection.cursor()
    cursor.execute("TRUNCATE TABLE  excelapp_exceluser")
    return HttpResponse("删除数据成功")
