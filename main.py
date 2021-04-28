from time import time
from classes import GetVehicleNos, LoginAndSearch, OutPut

# 程序开始计时
start = time()
# 全国查询-营运车辆
address = "http://10.100.32.31:8138/SQS/SQS/communal/iframe/frameProvinceVehicle.html"
target_mark = "预处理"
# 账号
# username = "陈雅明"
username = input("请输入你的运政账号：")
# 密码
# password = "123456"
password = input("请输入你的运政密码：")
# 获取车牌号码
get_vehicleNos = GetVehicleNos()
vehicleNos = get_vehicleNos.get_vehicleNos_by_xlsx(target_mark)
# 登录和查询
login_and_search = LoginAndSearch(address, username, password, vehicleNos)
[contents, particulars] = login_and_search.search_by_vehicleNos()
# 结果输出
output = OutPut(contents, particulars)
output.output_new_xls()
# 程序计时结束
end = time()
# 输出程序运行时间
print("\n", "处理结束，本次用时：{0:>6.2f}s".format(end - start))
input("提示:按“Enter“键结束本程序!")
