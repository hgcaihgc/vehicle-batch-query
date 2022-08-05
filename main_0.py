from time import time
from classes import GetVehicleNos, LoginAndSearch, OutPut

start = time()  # 程序开始计时
print("温馨提醒：本脚本为方便本单位人员批量查询车辆信息使用，请勿外泄，请勿作他用、滥用，进而造成资源浪费和带宽拥堵。")
print("温馨提醒：后续浏览器插件版正在筹划中，技术问题、环境部署问题、使用方法等请咨询作者（胡光潮，联系方式85331165）。")
print("="*100)
address = "http://10.100.32.31:8138/SQS/SQS/communal/iframe/frameProvinceVehicle.html"  # 全国查询-营运车辆 网址
target_mark = "预处理"  # 设置标记
username = input("请输入你的运政账号：")  # 输入账号
password = input("请输入你的运政密码：")  # 输入密码
plateColor = input("请输入车牌的颜色,车牌颜色关系到查询结果，请准确填写（蓝色、黄色、黑色、白色、黄绿色、渐变绿）其中一种：")  # 输入颜色
get_vehicleNos = GetVehicleNos()  # 获取车牌号码 实例化
vehicleNos = get_vehicleNos.get_vehicleNos_by_xlsx(target_mark)  # 获取车牌号
login_and_search = LoginAndSearch(address, username, password, vehicleNos, plateColor)  # 登录和查询 实例化
[contents, particulars] = login_and_search.search_by_vehicleNos()  # 查询
output = OutPut(contents, particulars)  # 输出 实例化
output.output_new_xls()  # 结果输出
end = time()  # 程序计时结束
print("="*100)
print("本次批量处理结束，共用时：{0:>6.2f}s".format(end - start))  # 输出程序运行时间
input("提示:按“Enter“键结束本程序!")


