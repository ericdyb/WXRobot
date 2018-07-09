# 读取Excel需求通报文件

import os
import sys
import getopt
import xlrd
from datetime import date,datetime
from pandas import DataFrame,Series
import string


DATA_DIR = 'd:/Python/WXRobot/data/'  # Excel数据文件存放目录
DATA_FILE = 'BI需求看板.xlsx'  # Excel数据文件存放目录

# SHEET1 = '重点关注临时需求（领导关注、集团上报需求、KPI需求）'


def read_excel():
    # 打开文件
    try:
        workbook = xlrd.open_workbook(os.path.join(DATA_DIR,DATA_FILE))
    except:
        print('读取EXCEL文件失败！')
        sys.exit(1)

    # 获取所有sheet
    print('已读取文件\'{}\',包含以下sheet: '.format(DATA_FILE))
    sheet_list = workbook.sheet_names()
    print(sheet_list)

    '''
    # 读取指定的sheet
    # sheet_focus = workbook.sheet_by_name(SHEET1)
    # print('已读取sheet\'{}\',包含{}行{}列。'.format(sheet_focus.name,sheet_focus.nrows,sheet_focus.ncols))

    # 读取指定的列
    col_reqprority = sheet_focus.col_values(0)
    col_reqid = sheet_focus.col_values(1)
    col_reqname = sheet_focus.col_values(2)
    col_reqdept = sheet_focus.col_values(5)
    col_reqfrom = sheet_focus.col_values(7)
    col_reqowner = sheet_focus.col_values(9)
    col_reqarrdate = sheet_focus.col_values(14)
    col_reqistoday = sheet_focus.col_values(17)
    col_reqforesee = sheet_focus.col_values(18)
    col_reqworker = sheet_focus.col_values(22)
    '''

    # 获取所有的sheet并组装写入数据框
    sheet_readin = []
    request_df = []
    # 第一张sheet是汇总说明不读取
    for j in range(workbook.nsheets):
        # 依次读取sheet
        sheet_readin.append(workbook.sheet_by_index(j))
        print('已读取sheet\'{}\',包含{}行{}列。'.format(sheet_readin[j].name, sheet_readin[j].nrows, sheet_readin[j].ncols))

        # 读取全部列
        dic_col = {}
        for i in range(sheet_readin[j].ncols):
            col_i = sheet_readin[j].col_values(i)
            # 将读取的多列组装为数据框
            dic_col[col_i[0]] = col_i[1:]

        # 将列组装成数据框
        request_df.append(DataFrame(dic_col))
        # print(request_df[j])

    return(request_df, sheet_list)

#按需求ID查询
def query_by_id(request_df, sheet_list, req_id):
    isfound = 0
    query_result = {}
    match_record = Series()
    for i in range(1,len(request_df)):
        if '需求单号' in request_df[i].columns:
            #match_record = request_df[i].loc[request_df[i].loc[:,'需求单号']==req_id]

            match_col = request_df[i].loc[:, '需求单号']
            for j in range(len(match_col)):
                if req_id in str(match_col[j]):
                    match_record = request_df[i].loc[j, :]
                    break

            if len(match_record) > 0:
                #print('【查询结果】需求ID为{}的记录位于第{}张表格\'{}\'：'.format(req_id,i+1,sheet_list[i]))

                #print('需求名称：', str(match_record['需求名称']).split(' ')[4].split('\n')[0])
                query_result['id'] = match_record['需求单号']
                query_result['name'] = match_record['需求名称']
                #print('提出部门：', str(match_record['提出部门']).split(' ')[4].split('\n')[0])
                query_result['dept'] = match_record['提出部门']
                #print('需求负责人：', str(match_record['需求负责人']).split(' ')[4].split('\n')[0])
                query_result['owner'] = match_record['需求负责人']
                #print('所处需求队列类别：', sheet_list[i])
                query_result['queue'] = sheet_list[i]
                #print('当前所处需求队列优先级：', int(str(match_record['序号']).split(' ')[4].split('\n')[0].split('.')[0]))
                query_result['priority'] = int(match_record['序号'])
                handler = match_record['处理人']
                if handler == '':
                    #print('当前处理人：暂无' )
                    query_result['handler'] = '暂无'
                else:
                    #print('当前处理人：',handler)
                    query_result['handler'] = handler
                try:
                    #print('完成进度：%.0f%%' % (100 * float(str(match_record['系统单完成状态']).split(' ')[4].split('\n')[0])))
                    query_result['status'] = 100 * float(match_record['系统单完成状态'])
                except:
                    #print('完成进度：暂无')
                    query_result['status'] = 0
                isfound = 1
                query_result['isfound'] = 1
                break
    if isfound == 0:
        #print('【查询结果】需求ID为{}的记录未找到。'.format(req_id))
        query_result['isfound'] = 0
    else:
        #print(query_result)
        pass

    return query_result


#按需求名称查询
def query_by_name(request_df, sheet_list, req_name):
    isfound = 0
    query_result = {}
    match_record = Series()
    for i in range(1,len(request_df)):
        if '需求名称' in request_df[i].columns:
            #match_record = request_df[i].loc[request_df[i].loc[:,'需求名称'].str.contains(req_name)]
            match_col = request_df[i].loc[:,'需求名称']
            for j in range(len(match_col)):
                if req_name in str(match_col[j]):
                    match_record = request_df[i].loc[j,:]
                    break

            if len(match_record) > 0:
                #print('【查询结果】需求ID为{}的记录位于第{}张表格\'{}\'：'.format(req_id,i+1,sheet_list[i]))

                #print('需求名称：', str(match_record['需求名称']).split(' ')[4].split('\n')[0])
                query_result['id'] =  match_record['需求单号']
                query_result['name'] = match_record['需求名称']
                #print('提出部门：', str(match_record['提出部门']).split(' ')[4].split('\n')[0])
                query_result['dept'] = match_record['提出部门']
                #print('需求负责人：', str(match_record['需求负责人']).split(' ')[4].split('\n')[0])
                query_result['owner'] = match_record['需求负责人']
                #print('所处需求队列类别：', sheet_list[i])
                query_result['queue'] = sheet_list[i]
                #print('当前所处需求队列优先级：', int(str(match_record['序号']).split(' ')[4].split('\n')[0].split('.')[0]))
                query_result['priority'] = int(match_record['序号'])
                handler = match_record['处理人']
                if handler == '':
                    #print('当前处理人：暂无' )
                    query_result['handler'] = '暂无'
                else:
                    #print('当前处理人：',handler)
                    query_result['handler'] = handler
                try:
                    #print('完成进度：%.0f%%' % (100 * float(str(match_record['系统单完成状态']).split(' ')[4].split('\n')[0])))
                    query_result['status'] = 100 * float(match_record['系统单完成状态'])
                except:
                    #print('完成进度：暂无')
                    query_result['status'] = 0
                isfound = 1
                query_result['isfound'] = 1
                break

    if isfound == 0:
        #print('【查询结果】需求名称为{}的记录未找到。'.format(req_name))
        query_result['isfound'] = 0
    else:
        #print(query_result)
        pass

    return query_result


#################################主函数###############################
# def request_query(argv):
def request_query(request_data, sheet_names, type, req):
    '''
    # 判断命令行调用规范性
    if argv.__len__() == 0:
        print('<Usage>', 'RequestQuery.py -i <需求ID> -n <需求名称>')
    try:
        opts, args = getopt.getopt(argv, "hi:n:")
    except getopt.GetoptError:
        print('<Usage>', 'RequestQuery.py -i <需求ID> -n <需求名称>')
        sys.exit(2)
    '''

    # 读取需求数据文件
    #request_data,sheet_names = read_excel()

    query_output = {}
    '''
    for opt, arg in opts:
        if opt == '-h':
            print('RequestQuery.py -i <需求ID> -n <需求名称>')
            sys.exit(0)
        elif opt == '-i':
            input_reqid = arg
            print('查询的需求ID：',input_reqid)
            query_output = query_by_id(request_data, sheet_names, input_reqid)
        elif opt == '-n':
            input_reqname = arg
            print('查询的需求名称：',input_reqname)
            query_output = query_by_name(request_data,sheet_names,input_reqname)
    '''
    query_output = {}
    if type == 'i':
        input_reqid = req
        print('查询的需求ID：', input_reqid)
        query_output = query_by_id(request_data, sheet_names, input_reqid)
    elif type == 'n':
        input_reqname = req
        print('查询的需求名称：', input_reqname)
        query_output = query_by_name(request_data, sheet_names, input_reqname)
    else:
        query_output['isfound'] = 0

    if query_output['isfound'] == 1:
        print('完成进度：%.0f%%' %query_output['status'])
    else:
        print('该需求查无记录，请检查输入信息是否正确。')

    return query_output

'''
if __name__ == '__main__':
   request_query(sys.argv[1:])
'''

