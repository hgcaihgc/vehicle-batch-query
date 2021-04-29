import os
import sys
import xlrd
import xlwt
import requests
from time import sleep
from datetime import datetime
from selenium import webdriver
from xlutils.copy import copy
from functions import get_vehicleNos_from_workbook, confirm_vehicleNo_index
from functions import search_by_vehicleNo, simplify_content, get_particulars


class GetVehicleNos:
    """功能模块类：获取车牌号码"""
    def get_vehicleNos_by_xlsx(self, target_mark="预处理"):
        """获取当前目录中所有目标文件workbook中的所有车牌号码"""        
        current_path = os.getcwd()  # 获取当前路径        
        file_list = os.listdir(current_path)  # 获取当前路径中的全部文件，形成列表        
        vehicleNos = []  # 定义车牌号码的空列表
        for file_name in file_list:            
            if target_mark in file_name:  # 如果文件名中有目标标记
                vehicleNos = vehicleNos + get_vehicleNos_from_workbook(file_name)
            else:
                pass
        if vehicleNos:
            return vehicleNos
        else:
            sys.exit("没有标记过的文件")


    def get_vehicleNos_by_picture(self):
        """获取当前目录中所有jpg格式文件的名称"""        
        current_path = os.getcwd()  # 获取当前路径        
        file_list = os.listdir(current_path)  # 获取当前路径中的全部文件，形成列表
        vehicleNos = []
        for file_name in file_list:
            if ".jpg" in file_name or ".JPG" in file_name:
                name = file_name.split(".")
                vehicleNos.append(name[0])
        return vehicleNos
  

class LoginAndSearch:
    """功能模块类：登录和搜索"""
    def __init__(self, address, username, password, vehicleNos, plateColor):        
        self.address = address
        self.username = username
        self.password = password
        self.vehicleNos = vehicleNos
        self.plateColor = plateColor

    def get_cookie(self):
        """模拟登陆，获取cookie"""
        print("开始登陆。")
        driver = webdriver.Chrome()  # 调用webdriver模块下的Chrome()类并赋值给变量driver        
        driver.implicitly_wait(5)  # 设置隐式等待时间5s        
        driver.get(self.address)  # 直接进入指定页面，自动跳转到登录页面        
        driver.find_element_by_id("username").send_keys(self.username)  # 输入账号        
        driver.find_element_by_id("password").send_keys(self.password)  # 输入密码
        sleep(0.5)  # 暂停0.5s
        driver.find_element_by_id("login-Button").click()  # 点击登录
        sleep(0.5)  # 暂停0.5s
        driver.find_element_by_tag_name("html").click()  # 切换到新的页面        
        cookies = driver.get_cookies()  # 获取新页面的cookie        
        cookie = 'JSESSIONID=' + cookies[0]['value']  # 合成可用的cookie        
        driver.quit()  # 退出浏览器
        print("登陆完成。")
        return cookie    

    def search_by_vehicleNos(self):
        """通过所有车牌号查询信息"""        
        cookie = self.get_cookie()  # 获取cookie信息
        print("开始查询。")
        vehicle_num = len(self.vehicleNos)  # 获取车牌号码列表的长度，即车辆的总数
        vehicle_new = self.vehicleNos[:]  # 复制车牌号列表
        contents = []  # 定义需要查询结果的空列表
        particulars = []  # 定义需要查询的详情的空列表
        message_progress = "正在查询{0:>8s}，这是该车第{1:>2d}次查询,进程为{2:>4d}/{3:>4d}, 进度为{4:>6.2f}%。"  # 定义进程信息
        i = 1
        total_count = 1
        while vehicle_new:
            count = 1  # 尝试连接的次数，初始化为0
            vehicleNo = vehicle_new[-1]
            while True:  #
                print('\r', message_progress.format(vehicleNo, count, i, vehicle_num, i*100/vehicle_num), end='', flush=True)  # 输出进程信息
                try:
                    content = search_by_vehicleNo(vehicleNo, cookie, self.plateColor)   # 通过车牌号码查询                    
                    content = simplify_content(content)  # 简化查询结果，主要是针对有读条记录的查询结果选择一个最新的
                    particular = get_particulars(content, cookie)  # 查询详情                    
                    contents.append(content)  # 将查询结果加入到结果列表
                    particulars.append(particular)  # 将详情结果加入到详情列表
                    i += 1
                    total_count += 1
                    vehicle_new.pop()
                    break  # 跳出该车牌号的查询
                except requests.exceptions.RequestException:  # 获取异常,查询异常时，如网速过慢超时
                    count += 1  # 查询次数加1
                    total_count += 1
                    sleep(0.5)
                    if count > 99:
                        sys.exit("网络不畅，请稍后再试！")
        print("\n", end='')
        print("查询结束。本次查询{0:>4d}辆车，共查询{1:>4d}次，平均每车查询{2:>4.2f}次。".format(vehicle_num, total_count, total_count/vehicle_num))
        return [contents, particulars]  # 返回查询结果、详情列表，完成查询和未完成查询的列表

    def search_by_vehicleNos_with_nonoperating(self):
        """从结果中筛选非营运的车辆"""
        contents = self.search_by_vehicleNos()
        contents_return = []
        contents_num = len(contents)
        message_progress = "正在筛选，进程为{0:>4d}/{1:>4d}, 进度为{2:>6.2f}%。"
        for i in range(contents_num):
            content = contents[i]            
            if content['total'] == 0:  # 没有营运信息的视为非营运
                contents_return.append(content)            
            else:                
                if content['rows'][0]['businessState'] != "营运":  # 营运状态显示不是“营运”的视为非营运
                    contents_return.append(content)                
                elif content['rows'][0]['certificateInfo'][0]['certificateExpireDate'] < datetime.now().strftime("%Y%m%d"):  # 营运状态为“营运”但是过期的视为非营运
                    contents_return.append(content)
                else:
                    pass
            # 输出进程信息
            print('\r', message_progress.format(i+1, contents_num, (i+1)*100/contents_num), end='', flush=True)
        return contents_return


class OutPut:
    """功能模块类：输出结果"""

    def __init__(self, contents, particulars):
        self.contents = contents
        self.particulars = particulars

    def output_new_xls(self):
        """将最后结果输出到xlsx文件"""
        print("开始写入")
        workbook = xlwt.Workbook()        
        sheet = workbook.add_sheet('结果', cell_overwrite_ok=True)  # 添加name为'结果'的sheet        
        title = ['车牌号码', '证件号码', '有效期至', '营运状态',
                 '经营范围', '发证机构', '业户名称', '经营许可证号',
                 '有效期至', '营运状态', '经营范围', '发证机构']  # 自定义标题行
        col_num = len(title)  # 获取列数
        [sheet.write(0, j, title[j]) for j in range(col_num)]  # 将标题行写到第1行
        contents_num = len(self.contents)  # 获取结果行数
        message_progress = "正在写入，进程为{0:>4d}/{1:>4d}, 进度为{2:>6.2f}%。"  # 定义进程信息
        for i in range(contents_num):
            content = self.contents[i]
            particular = self.particulars[i]            
            print('\r', message_progress.format(i+1, contents_num, (i+1)*100/contents_num), end='', flush=True)  # 输出进程信息
            if content['total'] > 0:
                for count in range(content['total']):                    
                    if 'certificateInfo' in content['rows'][count].keys():  # 在content有certificateInfo信息的情况时
                        if particular['total'] > 0:  # 如果particular有信息
                            if 'certificateInfo' in particular['rows'][0].keys():  # 在particular有certificateInfo信息的情况时
                                tr = [content['rows'][count]['vehicleNo'],
                                      content['rows'][count]['certificateInfo'][0]['transCertificateCode'],
                                      content['rows'][count]['certificateInfo'][0]['certificateExpireDate'],
                                      content['rows'][count]['businessState'],
                                      content['rows'][count]['certificateInfo'][0]['businessScopeDesc'],
                                      content['rows'][count]['certificateInfo'][0]['transCertificateGrantOrgan'],
                                      particular['rows'][0]['ownerName'],
                                      particular['rows'][0]['certificateInfo'][0]['licenseCode'],
                                      particular['rows'][0]['certificateInfo'][0]['expireDate'],
                                      particular['rows'][0]['operatingStatus'],
                                      particular['rows'][0]['certificateInfo'][0]['businessScopeDesc'],
                                      particular['rows'][0]['certificateInfo'][0]['licenseIssueOrgan']]
                            else:  # 在particular没有certificateInfo信息的情况时
                                tr = [content['rows'][count]['vehicleNo'],
                                      content['rows'][count]['certificateInfo'][0]['transCertificateCode'],
                                      content['rows'][count]['certificateInfo'][0]['certificateExpireDate'],
                                      content['rows'][count]['businessState'],
                                      content['rows'][count]['certificateInfo'][0]['businessScopeDesc'],
                                      content['rows'][count]['certificateInfo'][0]['transCertificateGrantOrgan'],
                                      particular['rows'][0]['ownerName'],
                                      'None',
                                      'None',
                                      particular['rows'][0]['operatingStatus'],
                                      'None',
                                      'None']
                        else:  # 如果particular没有信息
                            tr = [content['rows'][count]['vehicleNo'],
                                  content['rows'][count]['certificateInfo'][0]['transCertificateCode'],
                                  content['rows'][count]['certificateInfo'][0]['certificateExpireDate'],
                                  content['rows'][count]['businessState'],
                                  content['rows'][count]['certificateInfo'][0]['businessScopeDesc'],
                                  content['rows'][count]['certificateInfo'][0]['transCertificateGrantOrgan'],
                                  'None', 'None', 'None', 'None', 'None', 'None']              
                    else:  # 在content没有certificateInfo信息的情况时
                        tr = [content["vehicleNo"], 'None', 'None', 'None', 'None', 'None', 'None', 'None', 'None', 'None', 'None', 'None']
            else:  # 当content没有信息时
                tr = [content["vehicleNo"], 'None', 'None', 'None', 'None', 'None', 'None', 'None', 'None', 'None', 'None', 'None']
            [sheet.write(i+1, j, tr[j]) for j in range(col_num)]
        workbook.save('结果-{0}.xls'.format(datetime.now().strftime('%Y%m%d-%H%M%S')))
        print("\n", end='')
        print("写入完成。")


    def output_old_xls(self, target_mark="预处理"):
        """将最后结果输出到原xls文件，只有在只有一个xls文件时才适用"""
        print("开始写入")
        title = ['证件号码', '有效期至', '营运状态', '经营范围',
                 '发证机构', '业户名称', '经营许可证号', '有效期至',
                 '营运状态', '经营范围', '发证机构']  # 自定义标题行
        col_num = len(title)  # 获取列数        
        current_path = os.getcwd()  # 获取当前路径        
        file_list = os.listdir(current_path)  # 获取当前路径中的全部文件，形成列表
        for file_name in file_list:
            if target_mark in file_name:  # 如果文件名中有目标标记                
                work_book = xlrd.open_workbook(file_name)  # 打开xls表
                sheet = work_book.sheet_by_index(0)
                ncols = sheet.ncols
                index = confirm_vehicleNo_index(sheet)
            else:
                pass
        copy_work_book = copy(work_book)
        copy_sheet = copy_work_book.get_sheet(0)
        [copy_sheet.write(index[0], ncols+j, title[j]) for j in range(col_num)]
        contents_num = len(self.contents)
        for i in range(contents_num):
            # 输出进程信息
            print('\r', "正在写入，进程为{0:>4d}/{1:>4d}, 进度为{2:>6.2f}%".format(i+1, contents_num, (i+1)*100/contents_num), end='', flush=True)
            content = self.contents[i]
            particular = self.particulars[i]
            if content['total'] > 0:            
                for count in range(content['total']):
                    # 在有certificateInfo信息的情况时
                    if 'certificateInfo' in content['rows'][count].keys():
                        if particular['total'] > 0:
                            if 'certificateInfo' in particular['rows'][0].keys():  # 在particular有certificateInfo信息的情况时
                                tr = [content['rows'][count]['certificateInfo'][0]['transCertificateCode'],
                                    content['rows'][count]['certificateInfo'][0]['certificateExpireDate'],
                                    content['rows'][count]['businessState'],
                                    content['rows'][count]['certificateInfo'][0]['businessScopeDesc'],
                                    content['rows'][count]['certificateInfo'][0]['transCertificateGrantOrgan'],
                                    particular['rows'][0]['ownerName'],
                                    particular['rows'][0]['certificateInfo'][0]['licenseCode'],
                                    particular['rows'][0]['certificateInfo'][0]['expireDate'],
                                    particular['rows'][0]['operatingStatus'],
                                    particular['rows'][0]['certificateInfo'][0]['businessScopeDesc'],
                                    particular['rows'][0]['certificateInfo'][0]['licenseIssueOrgan']
                                    ]
                            else:  # 在particular没有certificateInfo信息的情况时
                                tr = [content['rows'][count]['vehicleNo'],
                                      content['rows'][count]['certificateInfo'][0]['transCertificateCode'],
                                      content['rows'][count]['certificateInfo'][0]['certificateExpireDate'],
                                      content['rows'][count]['businessState'],
                                      content['rows'][count]['certificateInfo'][0]['businessScopeDesc'],
                                      content['rows'][count]['certificateInfo'][0]['transCertificateGrantOrgan'],
                                      particular['rows'][0]['ownerName'],
                                      'None',
                                      'None',
                                      particular['rows'][0]['operatingStatus'],
                                      'None',
                                      'None']
                        else:
                            tr = [content['rows'][count]['vehicleNo'],
                                  content['rows'][count]['certificateInfo'][0]['transCertificateCode'],
                                  content['rows'][count]['certificateInfo'][0]['certificateExpireDate'],
                                  content['rows'][count]['businessState'],
                                  content['rows'][count]['certificateInfo'][0]['businessScopeDesc'],
                                  content['rows'][count]['certificateInfo'][0]['transCertificateGrantOrgan'],
                                  'None', 'None', 'None', 'None', 'None', 'None']
                    # 在没有certificateInfo信息的情况时
                    else:
                        tr = ['None', 'None', 'None', 'None', 'None', 'None', 'None', 'None', 'None', 'None', 'None']
            else:            
                tr = ['None', 'None', 'None', 'None', 'None', 'None', 'None', 'None', 'None', 'None', 'None']
            [copy_sheet.write(index[0]+1+i, ncols+j, tr[j]) for j in range(col_num)]
        copy_work_book.save(file_name)
        print("\n", end='')
        print("写入完成。")
