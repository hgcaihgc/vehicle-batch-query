import sys
import xlrd
import json
import requests



def get_vehicleNos(file_name):
    workbook = xlrd.open_workbook(file_name)
    names = workbook.sheet_names()
