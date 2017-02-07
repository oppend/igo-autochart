# !/usr/bin/env python
# coding=utf-8
"""
@name:整体汇总图表
@About:
@Author: oppend
@Contact: oppend@gmail.com
@SoftWare: PyCharm
@File: all_chart.py
@Time: 2017/01/16 23:00
# ============================单日整体汇总报表==============================
# 多部门多组织单日汇总报表，店长等查看
# ====================================================================
"""
import xlsxwriter
import time
import json

# 周数据字典读取
week_data_dic = json.load(open('./data/week_data.dic'))
# 部门字典读取
dept_code_dic = json.load(open('./config/dept_set.ini'))
# 组织代码读取
org_code_list = json.load(open('./config/org_set.ini'))
# 读取日数据字典文件
days_data_dic = json.load(open('./data/real_time_data.dic'))



# 获取当前时区时间
localtime = time.localtime()
today_is = str(localtime[2])
now_year = str(localtime[0])
now_month = str(localtime[1]).zfill(2)
now_day = str(localtime[2]).zfill(2)
now_hour = str(localtime[3])
now_minute = str(localtime[4])

# today_is = '18'
# now_month = '01'
# now_year = '2017'


# 定义数据表头列表
title = [u'部门', u'万达店', u'万达占比', u'世元店', u'世元占比', u'晨光店', u'晨光占比']


if __name__ == '__main__':
    # 创建一个Excel图表文件
    workbook = xlsxwriter.Workbook('./file/today_data.xlsx')
    # 创建一个工作表对象
    worksheet = workbook.add_worksheet()

    # 创建一个图表对象,设置图表类型:column柱状图 bar 横线图 pie饼图 stock股票样式
    chart1 = workbook.add_chart({'type': 'column'})
    # 定义部门名称format格式对象
    dept_name_format = workbook.add_format()
    # 定义format对象单元格边框加粗(1像素)的格式
    dept_name_format.set_border(1)
    # 定义format_title格式对象
    title_format = workbook.add_format()
    # 定义format_title对象单元格边框加粗(1像素)的格式
    title_format.set_border(1)
    # 定义format_title对象单元格背景颜色为'#cccccc'的格式
    title_format.set_bg_color('#cccccc')
    # 定义format_title对象单元格居中对齐的格式
    title_format.set_align('center')
    # 定义format_title对象单元格内容加粗的格式
    title_format.set_bold()

    # 定义数据图表样式
    chart1.add_series({
        'categories': '=Sheet1!$A$2:$A$17',
        'values': '=Sheet1!$B$2:$B$17',
        'name': '=Sheet1!$B$1',
        'line': {'color': '#FF9900'},
        'marker': {
            'type': 'diamond',
            'size': 8,
        },
        'data_labels': {'value': True},
        'smooth': True,
    })
    # 定义图表第2项数据
    chart1.add_series({
        'categories': ['Sheet1', 1, 0, 16, 0],  # 1行0列 到 7行0列
        'values': ['Sheet1', 1, 3, 16, 3],
        'name': ['Sheet1', 0, 3],
        'line': {'color': '#006633'},
        'marker': {
            'type': 'diamond',
            'size': 8,
            'fill': {'color': '#cc6600'},
        },
        'data_labels': {'value': True},
        'smooth': True,
    })
    # 定义图表第3项数据
    chart1.add_series({
        'categories': ['Sheet1', 1, 0, 16, 0],  # 1行0列 到 7行0列
        'values': ['Sheet1', 1, 5, 16, 5],
        'name': ['Sheet1', 0, 5],
        'line': {'color': '#CC3333'},
        'marker': {
            'type': 'diamond',
            'size': 8,
            'fill': {'color': '#0003366'},
        },
        'data_labels': {'value': True},
        'smooth': True,
    })
    # 图表标题设置
    chart1.set_title({
        'name': u'全店实时销售 %s-%s %s:%s' % (now_month, now_day, now_hour, now_minute),
        'font': {'size': 20, 'bold': 2},
        'layout': {
            'x': 0.42,
            'y': 0.05,
        }
    })
    # 设置图表图例
    chart1.set_legend({
        'font': {'size': 12, 'bold': 2},
    })
    # 设置图表数据表样式
    chart1.set_table({
        'show_keys': True,
        'font': {'size': 11, 'bold': 1}
    })

    # ====================表头===================
    # 写入表头
    # ==========================================

    # 部门名称列表初始化
    dept_name = []
    for dept_code in sorted(dept_code_dic.keys()):
        dept_name.append(dept_code_dic[dept_code])

    # 定义标题,部门名称,日期写入位置
    worksheet.write_row('A1', title, title_format)
    # 写入部门
    worksheet.write_column('A2', dept_name, dept_name_format)
    # 写入合计项
    worksheet.write_column('A18', [u'合计金额'], title_format)

    # ===================数据主体====================
    # 日数据字典格式｛部门：{日期:{组织:金额}}｝
    # ==============================================
    # 数据获取
    # 门店数据存储字典
    org_data_dic = dict()
    for org_code in org_code_list:
        org_data_dic[str(org_code)] = []
        org_tmp = []
        row_num = 1
        index_num = 0
        sum_total = 0
        for dept_code in sorted(dept_code_dic.keys()):
            dept_today_data = days_data_dic[dept_code]
            org_dept_data = dept_today_data[str(org_code)]
            org_tmp.append(org_dept_data)
            sum_total += float(str(org_dept_data))
            if str(org_code) == '1001':
                worksheet.write_number(row_num, 1, float(str(org_tmp[index_num])), dept_name_format)
                org_data_dic[str(org_code)].append(float(str(org_tmp[index_num])))
            elif str(org_code) == '1002':
                worksheet.write_number(row_num, 3, float(str(org_tmp[index_num])), dept_name_format)
                org_data_dic[str(org_code)].append(float(str(org_tmp[index_num])))
            else:
                worksheet.write_number(row_num, 5, float(str(org_tmp[index_num])), dept_name_format)
                org_data_dic[str(org_code)].append(float(str(org_tmp[index_num])))
            row_num += 1
            index_num += 1
        # 写入合计值
        if str(org_code) == '1001':
            worksheet.write_number(row_num, 1, float(str(sum_total)), dept_name_format)
            org_data_dic[str(org_code)].append(float(str(sum_total)))
        elif str(org_code) == '1002':
            worksheet.write_number(row_num, 3, float(str(sum_total)), dept_name_format)
            org_data_dic[str(org_code)].append(float(str(sum_total)))
        else:
            worksheet.write_number(row_num, 5, float(str(sum_total)), dept_name_format)
            org_data_dic[str(org_code)].append(float(str(sum_total)))

    # 设置图表样式皮肤
    chart1.set_style(66)
    # 设置图表大小
    chart1.set_size({'width': 1600, 'height': 900})
    # 图表写入位置
    worksheet.insert_chart('A22', chart1)

    # =========================部门占比========================
    # 创建一个图表对象,设置图表类型:column柱状图 bar 横线图 pie饼图 stock股票样式
    chart2 = workbook.add_chart({'type': 'pie'})
    for org_code in org_code_list:
        org_data = org_data_dic[str(org_code)]
        start_rows = 1
        total_rows = len(org_data)
        if str(org_code) == '1001':
            data_cols = 2
        elif str(org_code) == '1002':
            data_cols = 4
        else:
            data_cols = 6
        # 写入
        while start_rows <= total_rows:
            l_num = float(str(org_data[(start_rows - 1)]))
            r_num = float(str(org_data[total_rows - 1]))
            proportion = '%.2f%%' % (l_num/r_num*100)
            worksheet.write_string(start_rows, data_cols, proportion, dept_name_format)
            start_rows += 1

    # =====================支付方式占比=========================

    # 保存关闭文件
    workbook.close()
