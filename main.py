from time import time
from classes import GetVehicleNos, LoginAndSearch, OutPut

start = time()  # 程序开始计时
address = "http://10.100.32.31:8138/SQS/SQS/communal/iframe/frameProvinceVehicle.html"  # 全国查询-营运车辆 网址
target_mark = "预处理"  # 设置标记
username = input("请输入你的运政账号：")  # 输入账号
password = input("请输入你的运政密码：")  # 输入密码
get_vehicleNos = GetVehicleNos()  # 获取车牌号码 实例化
vehicleNos = get_vehicleNos.get_vehicleNos_by_xlsx(target_mark)  # 获取车牌号
login_and_search = LoginAndSearch(address, username, password, vehicleNos)  # 登录和查询 实例化
[contents, particulars] = login_and_search.search_by_vehicleNos()  # 查询
output = OutPut(contents, particulars)  # 输出 实例化
output.output_new_xls()  # 结果输出
end = time()  # 程序计时结束
print("本次批量查询结束，共用时：{0:>6.2f}s".format(end - start))  # 输出程序运行时间
input("提示:按“Enter“键结束本程序!")
