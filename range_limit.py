import xlrd
from xlutils.copy import copy
from openpyxl import *


def read_range_from_file_4(path):
    data_file_4 = xlrd.open_workbook(path)
    sheets = data_file_4.sheet_names()  # 获取工作簿中的所有表格
    worksheet = data_file_4.sheet_by_name(sheets[0])  # 获取工作簿中所有表格中的的第一个表格
    rows_old = worksheet.nrows  # 获取表格中已存在的数据的行数
    new_workbook = copy(data_file_4)  # 将xlrd对象拷贝转化为xlwt对象
    new_worksheet = new_workbook.get_sheet(0)  # 获取转化后工作簿中的第一个表格

    weihao_range_dict = {}
    target_range_arr = []
    for i in range(0,worksheet.nrows):
        if i == 0:
            continue
        else:
            range_str = worksheet.cell_value(i,3)
            left_rigth_value = range_str.replace("(","").replace(")","").replace("（","").replace("）","").split("-")
            range_arr = []
            if len(left_rigth_value) > 2:
                has_negative_value = False
                for j in range(0,len(left_rigth_value)):
                    if left_rigth_value[j] == "":
                        range_arr.append(-float(left_rigth_value[j+1]))
                        has_negative_value = True
                    else:
                        if has_negative_value:
                            has_negative_value = False
                            continue
                        else:
                            range_arr.append(float(left_rigth_value[j]))
            else:
                for j in range(0,len(left_rigth_value)):
                    range_arr.append(float(left_rigth_value[j]))
            weihao_range_dict[worksheet.cell_value(i, 1)] = range_arr

    print(weihao_range_dict)
    return weihao_range_dict



def filter_according_to_range(path,weihao_range_dict):
    data_file = xlrd.open_workbook(path)
    sheets = data_file.sheet_names()
    worksheet = data_file.sheet_by_name(sheets[0])
    rows_old = worksheet.nrows  # 获取表格中已存在的数据的行数
    # new_workbook = copy(data_file)  # 将xlrd对象拷贝转化为xlwt对象
    # new_worksheet = new_workbook.get_sheet(0)  # 获取转化后工作簿中的第一个表格


    cols_to_be_omit = []
    cols_name_to_be_omit = []
    # 处理附件1
    for i in range(16,worksheet.ncols):
        col_weihao_name = worksheet.cell_value(1,i)
        range_current_col_weihao_name = weihao_range_dict[col_weihao_name]
        for j in range(3,worksheet.nrows):
            value = worksheet.cell_value(j,i)
            range_left = range_current_col_weihao_name[0]
            range_right = range_current_col_weihao_name[1]
            if value>range_right or value<range_left:
                cols_to_be_omit.append({j:i})
                cols_name_to_be_omit.append({worksheet.cell_value(1,i):worksheet.cell_value(j,i)})
                break

    print(cols_to_be_omit)
    print(cols_name_to_be_omit)


def delete_cols_rows(path, cols_to_be_omit,rows_to_be_omit):
    wb = load_workbook(path)
    ws = wb.active
    ws.delete_cols(cols_to_be_omit)
    ws.delete_rows(rows_to_be_omit)  
    wb.save(path)



if __name__ == "__main__":
    weihao_range_dict = read_range_from_file_4("/Users/kuhn/Documents/Kuhn/PycharmProjects/HUAWEIB/data/附件四：354个操作变量信息.xlsx")
    filter_according_to_range("/Users/kuhn/Documents/Kuhn/PycharmProjects/HUAWEIB/data/附件一：325个样本数据.xlsx",weihao_range_dict)