import sys

import pandas as pd
from pyproj import Proj


class ReadExcel(object):
    def __init__(self, excel_file):
        self.excel_file: str = excel_file

    def read_calculate(self):
        # 读取
        xlsx = pd.ExcelFile(self.excel_file)
        sheet_values = {}
        for sheet_name in xlsx.sheet_names:
            sys.stdout.write("read sheet and calculate: " + sheet_name + "\n")
            content = pd.read_excel(xlsx, sheet_name)
            values = content.values
            cal_values = []
            for value in values:
                # 计算单行
                cal_value = self.calculate_line(value)
                cal_values.append(cal_value)
            sheet_values[sheet_name] = cal_values

        # 写出到新excel
        writer_file = self.excel_file.replace(".xlsx", "_calculated.xlsx")
        sys.stdout.write("write to new excel: " + writer_file + "\n")
        writer = pd.ExcelWriter(writer_file)
        for sheet_name in xlsx.sheet_names:
            content = pd.read_excel(xlsx, sheet_name)
            columns = content.columns.values
            df = pd.DataFrame(data=sheet_values[sheet_name], columns=columns)
            df.to_excel(writer, sheet_name, index=0)
        writer.save()
        writer.close()
        sys.stdout.write("write success\n")

    @staticmethod
    def calculate_line(value: list):
        # 原始维度
        x = value[1]
        # 原始经度
        y = value[2]
        if str(x) == "nan":
            return value
        # zone: 分带号； ellps：参考椭球
        # 默认为北半球 (north=True)，若坐标应在南半球添加south参数，使其为True即可
        p = Proj(proj="utm", zone=51, ellps="WGS84")
        lon, lat = p(x, y, inverse=True)  # 获取经纬度
        lon, lat = round(lon, 6), round(lat, 6)  # 经纬度保留6位小数
        # 将经纬度连接到字典的value中
        # 连接后字典value格式：[X, Y, lon, lat]
        value[3] = lon
        value[4] = lat
        return value
