import sys
import xlrd
import json
import requests


def confirm_index(sheet, names):
    """确定指定字段所在列的索引号"""
    # sheet的行数
    nrows = sheet.nrows
    # sheet的列数
    ncols = sheet.ncols
    # 初始化二维索引
    index = [None, None]
    for i in range(nrows):
        for j in range(ncols):
            # 按索引取sheet中对应的单元格的值，并与目标名称列表比对
            if sheet.cell_value(i, j) in names:
                index = [i, j]
                break
            else:
                pass
        if index != [None, None]:
            break
        else:
            pass
    else:
        pass
    return index


def confirm_vehicleNo_index(sheet):
    """确定车牌号码所在列的索引号"""
    # 定义车牌号码的可能名称的列表
    names = ["车牌号码", "号牌号码", "号码号牌",
             "车牌号", "车号码", "号码牌", "号牌码",
             "车牌", "车号", "号码", "号牌", "牌号"]
    # 找到车牌号码的索引号
    index = confirm_index(sheet, names)
    if index == [None, None]:
        sys.exit("没有指定的'车牌号码'字段")
    else:
        return index


def get_vehicleNos_from_sheet(sheet):
    """从一个sheet表中获取车牌号码"""
    index = confirm_vehicleNo_index(sheet)
    vehicleNos = sheet.col_values(index[1], start_rowx=index[0]+1, end_rowx=None) if index != [None, None] else []
    return vehicleNos


def get_vehicleNos_from_workbook(file_name, sheet_index="all"):
    """从一个workbook中获取所有sheet表中的车牌号码"""
    vehicleNos = []
    # 打开xlsx或者xls表
    workbook = xlrd.open_workbook(file_name)
    names = workbook.sheet_names()
    if sheet_index == "all" or sheet_index >= len(names):
        for name in names:
            sheet = workbook.sheet_by_name(name)
            vehicleNos = vehicleNos + get_vehicleNos_from_sheet(sheet)
    else:
        sheet = workbook.sheet_by_index(sheet_index)
        vehicleNos = get_vehicleNos_from_sheet(sheet)
    return vehicleNos


def get_provinceCode(vehicleNo):
    province_dict = {
        "京": "110000", "津": "120000", "冀": "130000", "晋": "140000", "蒙": "150000",
        "辽": "210000", "吉": "220000", "黑": "230000",
        "沪": "310000", "苏": "320000", "浙": "330000", "皖": "340000", "闽": "350000", "赣": "360000", "鲁": "370000",
        "豫": "410000", "鄂": "420000", "湘": "430000", "粤": "440000", "桂": "450000", "琼": "460000",
        "渝": "500000", "川": "510000", "贵": "520000", "云": "530000", "藏": "540000",
        "陕": "610000", "甘": "620000", "青": "630000", "宁": "640000", "新": "650000", "兵": "660000",
        "台": "710000", "香": "810000", "澳": "820000",
    }
    keys = list(province_dict.keys())
    acronym = vehicleNo[0]
    if acronym in keys:
        provinceCode = province_dict[acronym]
    else:
        provinceCode = "000000"
        print("{0}这辆车没有省份缩写码".format(vehicleNo))
    return provinceCode


def search_by_vehicleNo(vehicleNo, cookie, plateColor):
    """通过一个车牌号查询信息"""
    url = 'http://10.100.32.31:8138/SQS/overProvince/getResDirectory'
    plateColor_list = ['所有', '蓝色', '黄色', '黑色', '白色', '黄绿色', '渐变绿', '其他']
    plateColorCode = str(plateColor_list.index(plateColor))
    params_dict = {"vehicleNo": vehicleNo, "plateColorCode": plateColorCode, "provinceCode": get_provinceCode(vehicleNo)}
    From_data={
        'params': str(params_dict),
        'flag': 'vehicleByPlate',
        'page': '1',
        'rows': '10'
        }
    headers = {
        'Accept': 'application/json, text/javascript, */*; q=0.01',
        'Accept-Encoding': 'gzip, deflate',
        'Accept-Language': 'zh-CN,zh;q=0.9',
        'Connection': 'keep-alive',
        'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
        'Cookie': cookie,
        'Host': '10.100.32.31:8138',
        'Origin': 'http://10.100.32.31:8138',
        'Referer': 'http://10.100.32.31:8138/SQS/SQS/communal/iframe/frameProvinceVehicle.html',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/81.0.4044.113 Safari/537.36',
        'X-Requested-With': 'XMLHttpRequest'
        }
    # 获取请求的响应
    r = requests.post(url, data=From_data, headers=headers, timeout=1)
    # 获取响应中的json内容
    content = json.loads(r.text)
    # 在响应中添加车牌号码
    content["vehicleNo"] = vehicleNo
    return content


def get_particulars(content, cookie):
    """通过一个车牌号查询信息后再获取详情"""
    if content['total'] > 0:
        ownerId = content['rows'][0]['ownerId']
        provinceCode = content['rows'][0]['provinceCode']
        url = 'http://10.100.32.31:8138/SQS/overProvince/getResDirectory'
        From_data={
            'params': '{"ownerId":"' + ownerId + '","ownerName":"","provinceCode":"' + provinceCode + '"}',
            'flag': 'ownerByName'
            }
        headers = {
            'Accept': 'application/json, text/javascript, */*; q=0.01',
            'Accept-Encoding': 'gzip, deflate',
            'Accept-Language': 'zh-CN,zh;q=0.9',
            'Connection': 'keep-alive',
            'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
            'Cookie': cookie,
            'Host': '10.100.32.31:8138',
            'Origin': 'http://10.100.32.31:8138',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/81.0.4044.138 Safari/537.36',
            'X-Requested-With': 'XMLHttpRequest'
            }
        # 获取请求的响应
        r = requests.post(url, data=From_data, headers=headers, timeout=1)
        # 获取响应中的json内容
        particulars = json.loads(r.text)
        return particulars
    else:
        particulars = {"total": 0, "rows": [], "errorMsg": "交通部服务无法查询到指定数据"}
        return particulars


def simplify_content(content):
    """简化通过一个车牌号查询所得的信息，即从多条信息中筛选出最有用的"""
    # 没有营运登记信息的情况
    if content['total'] == 0:
        content_return = content
    # 只有1条营运登记信息的情况
    elif content['total'] == 1:
        content_return = content
    else:
        # 状态为营运的索引号列表
        operating_state_count = []
        # 有效期止信息的列表
        certificateExpireDate = []
        # 状态为营运的有效期止信息的列表
        operating_state_certificateExpireDate = []
        content_return = {}
        for count in range(content['total']):
            if 'certificateInfo' in content['rows'][count].keys():
                certificateExpireDate.append(content['rows'][count]['certificateInfo'][0]['certificateExpireDate'])
                if content['rows'][count]['businessState'] == "营运":
                    operating_state_certificateExpireDate.append(content['rows'][count]['certificateInfo'][0]['certificateExpireDate'])
                    operating_state_count.append(count)
                else:
                    pass
            else:
                pass
        # 状态为营运的信息条数为0
        if len(operating_state_count) == 0:
            # 有效期止时间最迟的索引号
            count = certificateExpireDate.index(max(certificateExpireDate))
        # 状态为营运的信息条数为1
        elif len(operating_state_count) == 1:
            # 索引号
            count = operating_state_count[0]
        else:
            # 状态为营运的有效期止时间最迟的索引号
            count = operating_state_certificateExpireDate.index(max(operating_state_certificateExpireDate))
        content_return["total"] = 1
        content_return["rows"] = [content['rows'][count]]
        content_return["errorMsg"] = content["errorMsg"]
        content_return["vehicleNo"] = content["vehicleNo"]
    return content_return
