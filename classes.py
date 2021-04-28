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
        # 获取当前路径
        current_path = os.getcwd()
        # 获取当前路径中的全部文件，形成列表
        file_list = os.listdir(current_path)
        # 定义车牌号码的空列表
        vehicleNos = []
        for file_name in file_list:
            # 如果文件名中有目标标记
            if target_mark in file_name:
                vehicleNos = vehicleNos + get_vehicleNos_from_workbook(file_name)
            else:
                pass
        if vehicleNos:
            return vehicleNos
        else:
            sys.exit("没有标记过的文件")


    def get_vehicleNos_by_picture(self):
        """获取当前目录中所有jpg格式文件的名称"""
        # 获取当前路径
        current_path = os.getcwd()
        # 获取当前路径中的全部文件，形成列表
        file_list = os.listdir(current_path)
        vehicleNos = []
        for file_name in file_list:
            if ".jpg" in file_name or ".JPG" in file_name:
                name = file_name.split(".")
                vehicleNos.append(name[0])
        return vehicleNos
  

class LoginAndSearch:
    """功能模块类：登录和搜索"""
    def __init__(self, address, username, password, vehicleNos):        
        self.address = address
        self.username = username
        self.password = password
        self.vehicleNos = vehicleNos

    def get_cookie(self):
        """模拟登陆，获取cookie"""
        # 调用webdriver模块下的Chrome()类并赋值给变量driver
        driver = webdriver.Chrome()
        # 设置隐式等待时间5s
        driver.implicitly_wait(5)
        # 直接进入指定页面，自动跳转到登录页面
        driver.get(self.address)
        # 输入账号
        driver.find_element_by_id("username").send_keys(self.username)
        # 输入密码
        driver.find_element_by_id("password").send_keys(self.password)
        sleep(0.5)
        # 点击登录
        driver.find_element_by_id("login-Button").click()
        sleep(0.5)
        # 切换到新的页面
        driver.find_element_by_tag_name("html").click()
        # 获取新页面的cookie
        cookies = driver.get_cookies()
        # 合成可用的cookie
        cookie = 'JSESSIONID=' + cookies[0]['value']
        # 退出浏览器
        driver.quit()
        return cookie    

    def search_by_vehicleNos(self):
        """通过所有车牌号查询信息"""
        cookie = self.get_cookie()
        # 获取车牌号码列表的长度，即车辆的总数
        vehicle_num = len(self.vehicleNos)
        # 定义需要查询结果的空列表
        contents = []
        particulars = []
        vehicleNos_finished = []
        vehicleNos_unfinished = []
        # 尝试连接的最大次数
        count_max = 10
        for i in range(vehicle_num):
            # 尝试连接的次数，初始化为0
            count = 0
            # 提取第i个车牌号码
            vehicleNo = self.vehicleNos[i]
            while count < count_max:
                try:
                    # 通过车牌号码查询
                    content = search_by_vehicleNo(vehicleNo, cookie)
                    # 简化查询结果
                    content = simplify_content(content)
                    particular = get_particulars(content, cookie)
                    vehicleNos_finished.append(vehicleNo)
                    break
                # 获取异常
                except requests.exceptions.RequestException:
                    count = count + 1
            else:
                print("完犊子了！{0}已经尝试{1:>3d}次连接了".format(vehicleNo, count_max))
                content = {'total': 0}
                particular = {"total": 0, "rows": [], "errorMsg": "交通部服务无法查询到指定数据"}
                # 在响应中添加车牌号码
                content["vehicleNo"] = vehicleNo                
                vehicleNos_unfinished.append(vehicleNo)
            contents.append(content)
            particulars.append(particular)
            # 输出进程信息
            print('\r', "正在查询{0}，进程为{1:>3d}/{2:>3d}, 进度为{3:>6.2f}%".format(vehicleNo, i+1, vehicle_num, (i+1)*100/vehicle_num), end='', flush=True)
        print("\n")
        return [contents, particulars, vehicleNos_finished, vehicleNos_unfinished]

    def search_by_vehicleNos_with_nonoperating(self):
        """从结果中筛选非营运的车辆"""
        contents = self.search_by_vehicleNos()
        contents_return = []
        contents_num = len(contents)
        for i in range(contents_num):
            content = contents[i]
            # 没有营运信息的视为非营运
            if content['total'] == 0:
                contents_return.append(content)            
            else:
                # 营运状态显示不是“营运”的视为非营运
                if content['rows'][0]['businessState'] != "营运":
                    contents_return.append(content)
                # 营运状态为“营运”但是过期的视为非营运
                elif content['rows'][0]['certificateInfo'][0]['certificateExpireDate'] < datetime.now().strftime("%Y%m%d"):
                    contents_return.append(content)
                else:
                    pass
            # 输出进程信息
            print('\r', "正在筛选，进程为{0:>3d}/{1:>3d}, 进度为{2:>6.2f}%".format(i+1, contents_num, (i+1)*100/contents_num), end='', flush=True)
        print("\n")
        return contents_return


class OutPut:
    """功能模块类：输出结果"""

    def __init__(self, contents, particulars):
        self.contents = contents
        self.particulars = particulars

    def output_new_xls(self):
        """将最后结果输出到xlsx文件"""
        workbook = xlwt.Workbook()
        # 添加name为'结果'的sheet
        sheet = workbook.add_sheet('结果', cell_overwrite_ok=True)
        # 自定义标题行
        title = ['车牌号码', '证件号码', '有效期至', '营运状态', '经营范围', '发证机构', '业户名称', '经营许可证号', '有效期至', '营运状态', '经营范围', '发证机构']
        col_num = len(title)
        [sheet.write(0, j, title[j]) for j in range(col_num)]
        contents_num = len(self.contents)
        for i in range(contents_num):
            content = self.contents[i]
            particular = self.particulars[i]
            if content['total'] > 0:
                for count in range(content['total']):
                    # 在有certificateInfo信息的情况时
                    if 'certificateInfo' in content['rows'][count].keys():
                        if particular['total'] > 0:
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
                                  particular['rows'][0]['certificateInfo'][0]['licenseIssueOrgan']
                                  ]
                        else:
                            tr = [content['rows'][count]['vehicleNo'],
                                  content['rows'][count]['certificateInfo'][0]['transCertificateCode'],
                                  content['rows'][count]['certificateInfo'][0]['certificateExpireDate'],
                                  content['rows'][count]['businessState'],
                                  content['rows'][count]['certificateInfo'][0]['businessScopeDesc'],
                                  content['rows'][count]['certificateInfo'][0]['transCertificateGrantOrgan'],
                                  'None', 'None', 'None', 'None', 'None', 'None']
                        [sheet.write(i+1, j, tr[j]) for j in range(col_num)]
                    # 在没有certificateInfo信息的情况时
                    else:
                        tr = [content["vehicleNo"], 'None', 'None', 'None', 'None', 'None', 'None', 'None', 'None', 'None', 'None', 'None']
                        [sheet.write(i+1, j, tr[j]) for j in range(col_num)]
            else:            
                tr = [content["vehicleNo"], 'None', 'None', 'None', 'None', 'None', 'None', 'None', 'None', 'None', 'None', 'None']
                [sheet.write(i+1, j, tr[j]) for j in range(col_num)]
            # 输出进程信息
            print('\r', "正在写入，进程为{0:>3d}/{1:>3d}, 进度为{2:>6.2f}%".format(i + 1, contents_num, (i + 1) * 100 / contents_num), end='', flush=True)
        print("\n")
        workbook.save('结果-{0}.xls'.format(datetime.now().strftime('%Y%m%d-%H%M%S')))

    def output_old_xls(self, target_mark="预处理"):
        """将最后结果输出到原xls文件，只有在只有一个xls文件时才适用"""
        # 自定义标题行
        title = ['证件号码', '有效期至', '营运状态', '经营范围', '发证机构', '业户名称', '经营许可证号', '有效期至', '营运状态', '经营范围', '发证机构']
        col_num = len(title)
        # 获取当前路径
        current_path = os.getcwd()
        # 获取当前路径中的全部文件，形成列表
        file_list = os.listdir(current_path)
        for file_name in file_list:
            # 如果文件名中有目标标记
            if target_mark in file_name:
                # 打开xls表
                work_book = xlrd.open_workbook(file_name)
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
            content = self.contents[i]
            particular = self.particulars[i]
            if content['total'] > 0:            
                for count in range(content['total']):
                    # 在有certificateInfo信息的情况时
                    if 'certificateInfo' in content['rows'][count].keys():
                        if particular['total'] > 0:
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
                        else:
                            tr = [content['rows'][count]['vehicleNo'],
                                  content['rows'][count]['certificateInfo'][0]['transCertificateCode'],
                                  content['rows'][count]['certificateInfo'][0]['certificateExpireDate'],
                                  content['rows'][count]['businessState'],
                                  content['rows'][count]['certificateInfo'][0]['businessScopeDesc'],
                                  content['rows'][count]['certificateInfo'][0]['transCertificateGrantOrgan'],
                                  'None', 'None', 'None', 'None', 'None', 'None']
                        [copy_sheet.write(index[0]+1+i, ncols+j, tr[j]) for j in range(col_num)]
                    # 在没有certificateInfo信息的情况时
                    else:
                        tr = ['None', 'None', 'None', 'None', 'None', 'None', 'None', 'None', 'None', 'None', 'None']
                        [copy_sheet.write(index[0]+1+i, ncols+j, tr[j]) for j in range(col_num)]
            else:            
                tr = ['None', 'None', 'None', 'None', 'None', 'None', 'None', 'None', 'None', 'None', 'None']
                [copy_sheet.write(index[0]+1+i, ncols+j, tr[j]) for j in range(col_num)]
            # 输出进程信息
            print('\r', "正在写入，进程为{0:>3d}/{1:>3d}, 进度为{2:>6.2f}%".format(i + 1, contents_num, (i + 1) * 100 / contents_num), end='', flush=True)
        print('\n')
        copy_work_book.save(file_name)
