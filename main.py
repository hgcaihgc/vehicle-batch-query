import requests
import xlrd
from xlutils.copy import copy
from time import sleep, time, localtime, strftime
import sys
from datetime import datetime


def get_student_information(id_number, cookie):
    """查询学员的报名信息"""
    url = 'http://10.145.149.223:8006/sxjgpt/student.do?list'
    headers = {
        'Accept': 'application/json, text/javascript, */*; q=0.01',
        'Accept-Encoding': 'gzip, deflate',
        'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6',
        'Connection': 'keep-alive',
        'Content-Length': '123',
        'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
        'Cookie': cookie,
        'Host': '10.145.149.223:8006',
        'Origin': 'http://10.145.149.223:8006',
        'Referer': 'http://10.145.149.223:8006/sxjgpt/student.do?main&menuId=402881e45aa6c5ae015aa6ca05cc0001',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/99.0.4844.74 Safari/537.36 Edg/99.0.1150.55',
        'X-Requested-With': 'XMLHttpRequest',
    }
    data = {
        'qCityCode': '3306',
        'qCountyCode': '',
        'qSchoolCode': '',
        'idnum': id_number,
        'stuName': '',
        'stuid': '',
        'phone': '',
        'newmodel': '',
        'traincar': '',
        'page': '1',
        'rows': '10',
    }
    # 获取请求的响应
    r = requests.post(url, data=data, headers=headers)
    response = r.json()
    if response['total'] == 0:
        applydate = '未报名'
        insName = '无'
    else:
        applydate = response['rows'][0]['applydate']
        insName = response['rows'][0]['insName']
    return applydate, insName


def get_training_record(id_number, cookie):
    """查询学员的阶段报审信息推送公安的情况"""
    url = 'http://10.145.149.223:8006/sxjgpt/stagetrainningtimeBak.do?list'
    headers = {
        'Accept': 'application/json, text/javascript, */*; q=0.01',
        'Accept-Encoding': 'gzip, deflate',
        'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6',
        'Connection': 'keep-alive',
        'Content-Length': '156',
        'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
        'Cookie': cookie,
        'Host': '10.145.149.223:8006',
        'Origin': 'http://10.145.149.223:8006',
        'Referer': 'http://10.145.149.223:8006/sxjgpt/stagetrainningtimeBak.do?main&menuId=8a11153e7aae6196017aae6399350002',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/99.0.4844.74 Safari/537.36 Edg/99.0.1150.55',
        'X-Requested-With': 'XMLHttpRequest',
    }
    data = {
        'qCityCode': '3306',
        'qCountyCode': '',
        'qSchoolCode': '',
        'subject': '',
        'auditstate': '-1',
        'applybegin': '',
        'applyend': '',
        'stuname': '',
        'idnum': id_number,
        'auditbegin': '',
        'auditend': '',
        'page': '1',
        'rows': '50'
    }
    # 获取请求的响应
    r = requests.post(url, data=data, headers=headers)
    response = r.json()
    return response


def stamp_to_str(time_stamp):
    # 转换本地时间
    # time_format = "%Y-%m-%d %H:%M:%S"
    time_format = "%Y-%m-%d"
    time1 = localtime(time_stamp)
    # 转为时间格式
    time2 = strftime(time_format, time1)
    return time2


def information_processing(response):
    """对学员的报审信息进行处理"""
    if response['total'] == 0:
        record = '未推送信息'
    else:
        record_model = '科目{0}报审时间：{1}；'
        record = ''
        for i in range(response['total']):
            record += record_model.format(response['rows'][i]['subject'], response['rows'][i]['crdate'])
    return record.rstrip()


def get_id_numbers_from_workbook(file_name):
    """从一个workbook中获取身份证信息"""    
    # 打开xlsx或者xls表
    workbook = xlrd.open_workbook(file_name)
    sheet = workbook.sheet_by_index(0)    
    id_numbers = sheet.col_values(1, start_rowx=1, end_rowx=None)
    return id_numbers


def get_information(file_name, cookie):
    """查询学员的信息，包括报名信息、阶段报审信息"""
    id_numbers = get_id_numbers_from_workbook(file_name)
    num = len(id_numbers)
    information = [['' for i in range(3)] for j in range(num)]
    for i in range(num):
        count = 1  # 尝试连接的次数，初始化为0
        top_count = 20
        id_number = id_numbers[i]
        message = "正在查询{0}，这是第{1:>2d}次查询，进程为{2:>4d}/{3:>4d}, 进度为{4:>6.2f}%"
        message_error = "{0}该名学员已查询超过{1}次，网络不畅，请稍后再试！"
        while True:
            print('\r', message.format(id_number, count, i+1, num, (i+1)*100/num), end='', flush=True)
            try:
                applydate, insName = get_student_information(id_number, cookie)
                information[i][0] = applydate
                information[i][1] = insName        
                information[i][2] = information_processing(get_training_record(id_number, cookie))
                break
            except requests.exceptions.RequestException:  # 获取异常,查询异常时，如网速过慢超时
                sleep(1)
                count += 1  # 查询次数加1
                if count > top_count:
                        sys.exit(message_error.format(id_number, top_count))
    return information


def output(file_name):
    information = get_information(file_name, cookie)
    title = ['报名时间', '培训机构', '信息推送公安情况']  # 自定义标题行
    col_num = len(title)  # 获取列数
    work_book = xlrd.open_workbook(file_name)  # 打开xls表
    sheet = work_book.sheet_by_index(0)
    ncols = sheet.ncols
    copy_work_book = copy(work_book)
    copy_sheet = copy_work_book.get_sheet(0)
    [copy_sheet.write(0, ncols+j, title[j]) for j in range(col_num)]
    [[copy_sheet.write(i+1, ncols+j, information[i][j]) for j in range(col_num)] for i in range(len(information))]
    copy_work_book.save('结果-{0}.xls'.format(datetime.now().strftime('%Y%m%d-%H%M%S')))


start = time()  # 程序开始计时
cookie = 'SECKEY_ABVK=NFCqxxfofanHUIdRUCEThekIRnVQpg1ZrFQLnb8Tez0%3D; BMAP_SECKEY=bBk3YN_47-la93E6EAIYTb3am2aRB9Ucdtu-_6J1IHYkJILB22AOG_lyYo2_UAkRWRqr7A0Q0afnbK0EytREyMgzwwslG-a1nE_9a78b4xZ9LIuPDAAr6NoFynM8sQKLwTlPmF9Y4GiqhOLAjVNm6c237xJza9VVtlp8B1SY2AlOe452oSh7UJkuhC3P3hx9; JSESSIONID=4E34A2AE65F1B6D687ABCBFA5BA92A1E'
file_name = '绍兴异地报名人员信息查询 - 预处理.xls'
output(file_name)
# b = get_id_numbers_from_workbook(file_name)
# print(b[0])
# print(type(b[0]))
# id_number = '230202198308101813'
# r = get_training_record(id_number, cookie)
# print(r["rows"][0]["stuName"] + str(r["rows"][0]["subject"]))
end = time()  # 程序计时结束
print('\n', "本次批量处理结束，共用时：{0:>6.2f}s".format(end - start))  # 输出程序运行时间
