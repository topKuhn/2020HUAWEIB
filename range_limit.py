import xlrd
from openpyxl import *
from xlutils.copy import copy
import pandas as pd
from pandas import DataFrame


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
        col_wei_hao_name = worksheet.cell_value(1,i)
        range_current_col_wei_hao_name = weihao_range_dict[col_wei_hao_name]
        for j in range(3,worksheet.nrows):
            value = worksheet.cell_value(j,i)
            range_left = range_current_col_wei_hao_name[0]
            range_right = range_current_col_wei_hao_name[1]
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


def get_range_from_file1(path):
    data_file_1 = xlrd.open_workbook(path)
    sheets = data_file_1.sheet_names()
    worksheet = data_file_1.sheet_by_name(sheets[0])
    wei_hao_range_dict = {}
    for i in range(16,worksheet.ncols):
        col_wei_hao_name = worksheet.cell_value(1,i)
        max_value = -10000.0
        min_value = 10000000.0
        for j in range(3,worksheet.nrows):
            cur_value = worksheet.cell_value(j,i)
            # if cur_value == 0:
            #     continue
            if cur_value > max_value:
                max_value = cur_value
            elif cur_value < min_value:
                min_value = cur_value
        cur_wei_hao_range = [min_value,max_value]
        wei_hao_range_dict[col_wei_hao_name]= cur_wei_hao_range
    print(wei_hao_range_dict)
    return wei_hao_range_dict


def filter_file_3_from_range(path, wei_hao_range_dict):
    data_file_3 = xlrd.open_workbook(path)
    sheets = data_file_3.sheet_names()
    worksheet = data_file_3.sheet_by_name(sheets[-1])
    rows_to_be_omit = []
    values_not_in_range = {}
    for i in range(42,82):
        for j in range(1,worksheet.ncols):
            cur_col_wei_hao_name = worksheet.cell_value(0,j)
            cur_value = worksheet.cell_value(i,j)
            cur_col_wei_hao_range = wei_hao_range_dict[cur_col_wei_hao_name]
            left_range = cur_col_wei_hao_range[0]
            right_range = cur_col_wei_hao_range[1]
            if (left_range-0.1) <= cur_value <= (right_range+0.1):
                continue
            else:
                rows_to_be_omit.append({i:cur_value})
                values_not_in_range[cur_col_wei_hao_name] = cur_value
                break

    print(rows_to_be_omit)
    print(values_not_in_range)

    return rows_to_be_omit


def cal_avg(path):
    data_file_3 = xlrd.open_workbook(path)
    sheets = data_file_3.sheet_names()
    cols_value = []
    print(data_file_3.sheet_by_name(sheets[1]).cell_value(1,3))
    minus = False
    for i in range(0,len(sheets)-1):
        worksheet = data_file_3.sheet_by_name(sheets[i])
        for j in range(3,worksheet.ncols):
            if i == 0:
                cols_value.append(worksheet.cell_value(1,j))
            elif i == 1:
                cols_value.append(worksheet.cell_value(2, j))
            elif i == 2:
                if not minus:
                    cols_value.append(cols_value[1]-cols_value[7])
                    minus = True
                cols_value.append(worksheet.cell_value(2, j))
            elif i == 3:
                cols_value.append(worksheet.cell_value(2, j))
    worksheet = data_file_3.sheet_by_name(sheets[-1])
    total_value = 0
    counts = 0
    for i in range(1,worksheet.ncols):
        for j in range(2,41):
            total_value = total_value + worksheet.cell_value(j,i)
            counts = counts + 1
        avg_value = total_value/counts
        cols_value.append(avg_value)


    print(cols_value)
    print(len(cols_value))

    return cols_value


def write_285_sample(path, cols_value):
    # data_file_1 = xlrd.open_workbook(path)
    # sheets = data_file_1.sheet_names()
    # worksheet = data_file_1.sheet_by_name(sheets[0])
    # new_data_file = copy(data_file_1)
    # new_worksheet = new_data_file.get_sheet(0)
    # print(new_worksheet.cell_value(287,0))
    # for i in range(2,worksheet.ncols):
    #     for j in range(0,len(cols_value)):
    #         new_worksheet.write(287,i,cols_value[j])
    # new_data_file.save(path)
    data = pd.read_excel(path,sheet_name="Sheet1")
    cols_value.insert(0,"2017/7/17 8:00:00")
    cols_value.insert(0,285)
    print(cols_value)
    # data = data.drop(287,axis = 0)
    data.loc[286] = cols_value

    DataFrame(data).to_excel('/Users/kuhn/Documents/Kuhn/PycharmProjects/HUAWEIB/data/test.xlsx',sheet_name="Sheet1",index=False,header=True)





if __name__ == "__main__":
    file_4 = "/Users/kuhn/Documents/Kuhn/PycharmProjects/HUAWEIB/data/附件四：354个操作变量信息.xlsx"
    file_1 = "/Users/kuhn/Documents/Kuhn/PycharmProjects/HUAWEIB/data/附件一：325个样本数据.xlsx"
    file_3 = "/Users/kuhn/Documents/Kuhn/PycharmProjects/HUAWEIB/data/附件三：285号和313号样本原始数据.xlsx"
    # wei_hao_range_dict = read_range_from_file_4(file_4)
    # filter_according_to_range(file_1,wei_hao_range_dict)

    wei_hao_range_dict = get_range_from_file1(file_1)
    rows_to_be_omit = filter_file_3_from_range(file_3,wei_hao_range_dict)

    print("sdfasdfasdf")
    cols_value = cal_avg(file_3)

    write_285_sample(file_1,cols_value)