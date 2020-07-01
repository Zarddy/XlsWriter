import time

import xlsxwriter

from helper.DatabaseHelper import DatabaseHelper


class ChartHelper:

    def __init__(self) -> None:
        super().__init__()
        self.db = DatabaseHelper()


    """
    导出设备故障记录，按单独故障类型分页
    """
    def export_device_fault_table(self):

        # 时间戳
        time_string = time.strftime('%Y%m%d_%H%M%S', time.localtime())
        filename = "故障记录表_%s.xlsx" % time_string

        # 创建一个excel
        workbook = xlsxwriter.Workbook(filename)

        name_array = ["SN码", "客户名称", "故障类型备注", "出厂日期", "使用时长（小时）",
                "故障发生日期", "故障摘要", "故障原因", "跟进人", "解决方法", "处理人", "处理结果", "备注"]

        # 获取所有故障类型
        category_list = self.db.select('select * from crm_device_info_category where type=2 and status=1')

        for category in category_list:
            # 标签名称
            sheet_name = category['category_name']

            # 创建一个sheet
            worksheet = workbook.add_worksheet(sheet_name)

            # 自定义样式，加粗
            header_format = workbook.add_format({'bold': 1})
            header_format.set_align('center')
            header_format.set_align('vcenter')

            # 日期格式
            date_format = workbook.add_format({'num_format': 'yyyy-mm-dd'})

            # 写入数据
            worksheet.write_row('A1', name_array, cell_format=header_format)  # 类型

            # 设置默认行高
            worksheet.set_default_row(height=18)
            # 设备单元格宽度
            worksheet.set_column(0, 0, 15)
            worksheet.set_column(1, 3, 13)
            worksheet.set_column(4, 4, 16)
            worksheet.set_column(6, 7, 30)
            worksheet.set_column(9, 9, 30)
            # 设置第5列的格式为日期
            worksheet.set_column(first_col=5, last_col=5, width=14, cell_format=date_format)

            # 通过类型Id查出对应的故障记录
            sql = 'SELECT device_sn, customer_name, fault_category_remark, production_time,' \
                  ' usage_time, fault_time, fault_summary, fault_reason, follow_up_person, resolvent, `handler`,' \
                  ' handling_result, remark from crm_device_fault_record where status=1 and fault_category_id=%d' % category['category_id']
            record_list = self.db.select(sql)

            # 写入记录
            for i in range( record_list.__len__() ):
                record = dict(record_list[i])
                v = record.values()
                worksheet.write_row('A' + str(i+2),  v)

        workbook.close()
        pass


    """
    导出设备故障记录，数据统计图表，按客户分页
    """
    def export_device_fault_chart(self):

        # 时间戳
        time_string = time.strftime('%Y%m%d_%H%M%S', time.localtime())
        filename = "故障记录统计图_%s.xlsx" % time_string

        # 创建一个excel
        workbook = xlsxwriter.Workbook(filename)

        # 获取客户列表
        sql_list_customers = 'select * from crm_customer where status = 1 order by customer_name asc'
        customer_list = self.db.select(sql_list_customers)

        last_one = customer_list[0]
        last_one['ids'] = [last_one['customer_id']]
        new_customer_list = []

        for customer in customer_list:
            if last_one['customer_id'] == customer['parent_id']:
                last_one['ids'] = last_one['ids'] + [customer['customer_id']]
            else:
                last_one = customer
                last_one['ids'] = [last_one['customer_id']]
                new_customer_list.append(last_one)

        for customer in new_customer_list:
            customer_name = customer['customer_name']
            sheet_name = customer_name

            ids = str(customer['ids'])[1:-1]  # id组合
            sql = "SELECT fault_category, count(record_id) as 'count' FROM hkphotonics_crm.crm_device_fault_record where customer_id in (" + ids + ") and status=1 group by fault_category"
            fault_category_list = self.db.select(sql)

            if fault_category_list.__len__() == 0:
                continue

            array_fault_category = []
            array_count = []
            for category in fault_category_list:
                array_fault_category.append(category['fault_category'])
                array_count.append(category['count'])

            # 创建一个sheet
            worksheet = workbook.add_worksheet(sheet_name)

            # 自定义样式，加粗
            bold = workbook.add_format({'bold': 1})

            # --------1、准备数据并写入excel---------------
            # 向excel中写入数据，建立图标时要用到
            headings = ['故障类型', '故障次数', '百分比']
            data = [
                array_fault_category,
                array_count,
            ]

            # 写入表头
            worksheet.write_column('A2', headings, bold)

            # 写入数据
            worksheet.write_row('B2', data[0])  # 类型
            worksheet.write_row('B3', data[1])  # 次数
            total = sum(data[1])  # 总次数

            percent_array = []
            for d in data[1]:
                percent_array.append('%.1f%%' % (d / total * 100))

            worksheet.write_row('B4', percent_array)  # 百分比

            # --------2、生成图表并插入到excel---------------
            # 创建一个柱状图(column chart)
            chart_col = workbook.add_chart({'type': 'column'})

            count = fault_category_list.__len__()

            color_list = ['#ff5555', '#61a8de', '#129b57', '#b9d042', '#fcc6bc', '#f7bf50',
                          '#1c7dca', '#8096cf', '#c0156c', '#47bba0', '#d8b24d', '#87ceeb']
            points = []
            for i in range(count):
                points.append({"fill": {"color": color_list[i]}})

            c = ord('B') + (count - 1)  # 总数量，在excel表中的最后一列数据的列序号ascii码的值

            # 数据集合
            chart_data_series = {
                # 这里的sheet1是默认的值，因为我们在新建sheet时没有指定sheet名
                # 如果我们新建sheet时设置了sheet名，这里就要设置成相应的值
                'name': sheet_name,
                'categories': '=' + sheet_name + '!$B$2:$%c$2' % c,
                'values': '=' + sheet_name + '!$B$3:$%c$3' % c,
                "points": points  # 定义各饼块的颜色
            }

            # 配置第一个系列数据
            chart_col.add_series(chart_data_series)

            # 设置图表的title 和 x，y轴信息
            chart_col.set_title({'name': '故障类型与故障次数'})
            chart_col.set_x_axis({'name': headings[0]})
            chart_col.set_y_axis({'name': headings[1]})

            # 设置图表的风格
            chart_col.set_style(1)

            # 把图表插入到worksheet以及偏移
            worksheet.insert_chart('A7', chart_col, {'x_offset': 25, 'y_offset': 10, 'x_scale': 1.2, 'y_scale': 1.3})

            # 添加饼图
            chart3 = workbook.add_chart({"type": "pie"})
            chart3.add_series(chart_data_series)
            chart3.set_title({"name": "故障类型百分比"})
            chart3.set_style(3)

            # 把图表插入到worksheet以及偏移
            worksheet.insert_chart('K7', chart3, {'x_offset': 25, 'y_offset': 10, 'x_scale': 1.2, 'y_scale': 1.3})

        workbook.close()
