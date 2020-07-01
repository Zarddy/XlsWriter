
from helper.ChartHelper import ChartHelper

if __name__ == '__main__':
    chart = ChartHelper()

    # 导出记录表，按单独故障类型分页
    chart.export_device_fault_table()

    # 导出统计图，按客户分页
    chart.export_device_fault_chart()
