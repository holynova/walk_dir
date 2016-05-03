# -*- coding: utf-8 -*-
#遍历文件夹取文件名
import os,os.path
import re
from openpyxl import Workbook,load_workbook

dir_name = u"E:\kuaipan\非投标任务\\2016年5月3日 micheal 要的投标记录统计\开标记录\\2014年开标汇总及分析表"
# dir_name = u"E:\\kuaipan\\非投标任务\\2016年5月3日 micheal 要的投标记录统计\\开标记录"
dir_name = u'E:\kuaipan\非投标任务\\2016年5月3日 micheal 要的投标记录统计\开标记录'

bid_arr = []
def walk_dir(dir,topdown = True):
    names = []
    for root,dirs,files in os.walk(dir,topdown):
        # for d in dirs:
            # print d
        for f in files:
            patten = re.compile(r'NCL-[TP]1\d-\d+')
            obj_match = patten.search(f)
            if obj_match:
                # print obj_match.group(0)
                full_name = os.path.join(root,f)
                # names.append(full_name)
                bid = Bid()
                bid.wb = full_name
                bid.
                find_worksheets(full_name)
   
    # return names

class Bid:
    def __init__(self,workbook="",worksheets=[]):
        self.wb = workbook
        self.shts = worksheets
    # pass


def find_worksheets(excel_file_name):
    patten = re.compile(r'.xlsx')
    if patten.search(excel_file_name):
        try:
            wb = load_workbook(excel_file_name)
        except :
            print "%s load failed" %(excel_file_name)
        else:
            print excel_file_name
            i = 0
            for sht in wb.get_sheet_names():
                i += 1
                print "   %d,%s" %(i,sht)
            # print wb.get_sheet_names()


names = walk_dir(dir_name) 


