from openpyxl import Workbook
from openpyxl.chart import (
    BarChart, LineChart, PieChart,
    Reference, Series
)
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.title import Title


class ExcelChartManager:
    """Excel图表管理类，支持创建、配置和插入多种图表"""
    
    def __init__(self, workbook=None):
        """初始化图表管理器"""
        self.workbook = workbook if workbook else Workbook()
        self.current_sheet = self.workbook.active  # 默认使用当前活跃工作表

    def set_sheet(self, sheet_name=None):
        """切换或创建目标工作表"""
        if sheet_name:
            if sheet_name in self.workbook.sheetnames:
                self.current_sheet = self.workbook[sheet_name]
            else:
                self.current_sheet = self.workbook.create_sheet(title=sheet_name)
        return self  # 支持链式调用
    
    @staticmethod
    def get_report_by_first_column(sheet, flag):
        """
        获取第一列的报告数据
        :param sheet: 工作表对象
        :param flag: 标志位，0表示直接对比1，1表示按索引列对比，2表示按索引列和映射列对比

        :return: 报告数据列表
        """
        # 存储第一列数据
        first_column = []
        
        # 遍历第一列的所有行（min_col=1, max_col=1 表示只取第一列）
        for row in sheet.iter_rows(min_col=1, max_col=1, values_only=True):
            # row是一个元组，包含当前行第一列的单元格值
            first_column.append(row[0])
            
        # 统计一致、差异点、新增、删除行的数量
        agreed_rows_number = sum(1 for item in first_column if item == "一致")
        not_agreed_rows_number = sum(1 for item in first_column if item == "差异点")
        unique_rows_number = sum(1 for item in first_column if item == "新增")
        delete_rows_number = sum(1 for item in first_column if item == "删除行")
        if flag == 0:
            data = [
                ["一致行数", agreed_rows_number     ],
                ["差异行数", not_agreed_rows_number ],
            ]
        else:
            data = [
                ["一致行数", agreed_rows_number     ],
                ["差异行数", not_agreed_rows_number ],
                ["新增行数", unique_rows_number     ],
                ["删除行数", delete_rows_number     ],
            ]

        return data

    def add_test_data(self, data: list):
        """添加测试数据（用于示例）"""
        
        for row in data:
            self.current_sheet.append(row)
        return self  # 支持链式调用

    def add_cell_value(self, row, column, value_list):
        """
        在指定行和列添加单元格内容
        :param row: 行号（从1开始）
        :param column: 列号（从1开始，1=A列，2=B列...）
        :param value: 要添加的值
        """
        print(f"row = {row}\n, column = {column}\n, value_list = {value_list}")

        rows = row
        for row_data in value_list:
            col = column
            for data in row_data:
                self.current_sheet.cell(row=rows, column=col, value=data)
                col += 1
            rows += 1
                
                
    def create_bar_chart(self, title, data_range, categories_range, pos="E2"):
        """创建柱状图"""
        chart = BarChart()
        chart.title = title  # 直接使用字符串作为标题（兼容所有版本）
        chart.style = 10  # 预设样式
        
        # 添加数据系列
        for col in range(data_range.min_col, data_range.max_col + 1):
            series_data = Reference(
                self.current_sheet,
                min_col=col,
                min_row=data_range.min_row,
                max_row=data_range.max_row
            )
            series = Series(series_data, title_from_data=True)
            chart.append(series)
        
        # 设置分类轴（X轴）
        chart.set_categories(categories_range)
        
        # 设置坐标轴标题（直接用字符串）
        chart.x_axis.title = "类别"
        chart.y_axis.title = "数值"
        
        # 插入图表
        self.current_sheet.add_chart(chart, pos)
        return self

    def create_line_chart(self, title, data_range, categories_range, pos="E18"):
        """创建折线图"""
        chart = LineChart()
        chart.title = title  # 直接使用字符串标题
        chart.style = 12
        chart.marker = True  # 显示数据点标记
        
        for col in range(data_range.min_col, data_range.max_col + 1):
            series_data = Reference(
                self.current_sheet,
                min_col=col,
                min_row=data_range.min_row,
                max_row=data_range.max_row
            )
            series = Series(series_data, title_from_data=True)
            chart.append(series)
        
        chart.set_categories(categories_range)
        chart.x_axis.title = "类别"
        chart.y_axis.title = "数值"
        
        self.current_sheet.add_chart(chart, pos)
        return self
    
    @staticmethod
    def create_referencec_data(sheet, labels_col, data_col, min_row, max_row):

        """创建饼图数据范围"""
        # 定义饼图标签范围
        pie_labels_range = Reference(
            sheet,
            min_col=labels_col,
            min_row=min_row,
            max_row=max_row
        )
        # 定义饼图数据范围
        pie_data_range = Reference(
            sheet,
            min_col=data_col,
            min_row=min_row,
            max_row=max_row
        )
        return pie_labels_range, pie_data_range

    def create_pie_chart(self, title, data_range, labels_range, pos="E34"):
        """创建饼图"""
        chart = PieChart()
        chart.title = title  # 直接使用字符串标题
        chart.style = 15
        
        # 添加数据系列
        series = Series(data_range, title_from_data=True)
        chart.series = [series]
        chart.set_categories(labels_range)
        
        # 显示数据标签（数值+百分比）
        chart.dataLabels = DataLabelList()
        chart.dataLabels.showVal = True
        chart.dataLabels.showPercent = True
        
        self.current_sheet.add_chart(chart, pos)
        return self

    def save(self, filename):
        """保存工作簿"""
        self.workbook.save(filename)
        print(f"文件已保存：{filename}")


def main():
    # 1. 初始化图表管理器
    chart_manager = ExcelChartManager()
    
    # 2. 创建并切换到目标工作表
    chart_manager.set_sheet("销售数据图表")
    
    # 3. 添加测试数据
    chart_manager.add_test_data()
    
    # 4. 定义数据范围（行列索引从1开始）
    data_range = Reference(
        chart_manager.current_sheet,
        min_col=2,  # B列
        min_row=2,  # 第2行
        max_col=4,  # D列
        max_row=6   # 第6行
    )
    categories_range = Reference(
        chart_manager.current_sheet,
        min_col=1,  # A列
        min_row=2,
        max_row=6
    )
    pie_labels_range = Reference(
        chart_manager.current_sheet,
        min_col=1,
        min_row=2,
        max_row=4
    )
    pie_data_range = Reference(
        chart_manager.current_sheet,
        min_col=2,
        min_row=2,
        max_row=4
    )
    
    # 5. 创建图表
    chart_manager.create_bar_chart(
        title="月度销售数据柱状图",
        data_range=data_range,
        categories_range=categories_range,
        pos="E2"
    ).create_line_chart(
        title="月度销售数据折线图",
        data_range=data_range,
        categories_range=categories_range,
        pos="E18"
    ).create_pie_chart(
        title="前3月销量占比饼图",
        data_range=pie_data_range,
        labels_range=pie_labels_range,
        pos="E34"
    )
    
    # 6. 保存文件
    chart_manager.save("销售数据图表示例.xlsx")

if __name__ == "__main__":
    main()