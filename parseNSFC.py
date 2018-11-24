#!/bin/sh/env python3
# -*- coding: utf-8 -*-
'''
Created on 2018-11-24
@author: Huang Jinwei
@description: 从isisn.nsfc.gov.cn获取多年的重大项目列表，并保存到csv文件，可选自动识别验证码
'''

from PIL import Image
from io import BytesIO
import requests as req
import os, sys, re
import time, datetime
import xml.dom.minidom as xmldom
import pytesseract

result_xml = 'result.xml'
result_csv = 'result.csv'
nsfc_req_url = "https://isisn.nsfc.gov.cn/egrantindex/funcindex/prjsearch-list?flag=grid&checkcode="
nsfc_req_raw_url = "https://isisn.nsfc.gov.cn/egrantindex/funcindex/prjsearch-list"
nsfc_validcode_img_url = "https://isisn.nsfc.gov.cn/egrantindex/validatecode.jpg"
nsfc_validcode_chk_url = "https://isisn.nsfc.gov.cn/egrantindex/funcindex/validate-checkcode"

def get_nsfc_data(year, result_xml, autoverify=False):

    req_session = req.session()
    headers = {'content-type': 'application/x-www-form-urlencoded; charset=UTF-8'}

    # get verify code image
    vimg_response = req_session.get(nsfc_validcode_img_url)
    image = Image.open(BytesIO(vimg_response.content))

    if autoverify == True:
        imgry = image.convert('L')  # 转化为灰度图
        text = pytesseract.image_to_string(imgry)
        print("validate code = "+text)
        verify_c_str = text
    else:
        image.show()
        verify_c_str = input("Please input verify code: ")

    validate_resp = req_session.post(nsfc_validcode_chk_url, data={"checkCode":verify_c_str})
    #print(verify_c_str)
    #print(validate_resp.text)

    if(validate_resp.text != "success"):
        print("Validate fail!")
    else:
        print("Validate success!")

    raw_req_data = {
        "resultDate": "prjNo:,ctitle:,psnName:,orgName:,subjectCode:,f_subjectCode_hideId:,subjectCode_hideName:,keyWords:,checkcode:"+verify_c_str+",grantCode:222,subGrantCode:,helpGrantCode:,year:" + str(year),
        "checkcode": verify_c_str
    }

    req_resp = req_session.post(nsfc_req_raw_url, data=raw_req_data, verify=False, headers=headers)
    #print(req_resp.text)

    req_data = {
        "_search": "false",
        "nd": str((int(round(time.time() * 1000)))),
        "rows": "10000",
        "page": "1",
        "sidx": "",
        "sord": "desc",
        "searchString": "resultDate^:prjNo%3A%2Cctitle%3A%2CpsnName%3A%2CorgName%3A%2CsubjectCode%3A%2Cf_subjectCode_hideId%3A%2CsubjectCode_hideName%3A%2CkeyWords%3A%2Ccheckcode%3A"+verify_c_str+"%2CgrantCode%3A222%2CsubGrantCode%3A%2ChelpGrantCode%3A%2Cyear%3A"+str(year)+"[tear]sort_name1^:psnName[tear]sort_name2^:prjNo[tear]sort_order^:desc"
    }

    req_resp = req_session.post(nsfc_req_url, data=req_data, verify=False, headers=headers)
    #print(req_resp.text)

    fh_xml = open (result_xml, 'w+', encoding='utf-8-sig')
    fh_xml.write ( req_resp.text ) 
    fh_xml.close()

def trans2csv(year, result_xml, result_csv, append=False, startnum=1):
    # parse the result xml
    # 得到文档对象
    domobj = xmldom.parse(result_xml)
    #print("xmldom.parse:", type(domobj))

    # 得到元素对象
    elementobj = domobj.documentElement
    #print ("domobj.documentElement:", type(elementobj))

    #获得子标签
    subElementObj = elementobj.getElementsByTagName("row")
    #print ("getElementsByTagName:", type(subElementObj))
    #print (len(subElementObj))

    open_mode = 'w+'

    if append == True:
        open_mode = 'a+'

    fh_csv = open (result_csv, open_mode, encoding='utf-8-sig') 

    wr_count = 0

    # 获得标签属性值
    for i in range(len(subElementObj)):
        cellObj = subElementObj[i].getElementsByTagName("cell")
        #print (len(cellObj))
        fh_csv.write(str(startnum + wr_count) + ',' + str(year) + ',')
        for i in range(len(cellObj)):
            fh_csv.write(cellObj[i].childNodes[0].data.replace(',','，').replace('&nbsp;', ' '))
            if i != len(cellObj) - 1:
                fh_csv.write(',')
            else:
                fh_csv.write('\n')
        wr_count += 1
    fh_csv.close()

    return len(subElementObj)


if __name__ == "__main__":
    req_year = [2014, 2015, 2016, 2017, 2018]
    csv_header = '序号,查询年份,项目批准号,申请代码1,项目名称,项目负责人,依托单位,批准金额,项目起止年月\n'

    if(os.path.exists(result_csv)):
        os.remove(result_csv)

    fh_csv = open (result_csv, 'w+', encoding='utf-8-sig')
    fh_csv.write(csv_header)
    fh_csv.close() 

    last_count = 0
    year_count = 0
    while year_count < len(req_year):
        year = req_year[year_count]
        try:
            cur_result_xml = str(year) + '_' + result_xml
            print("Parsing " + str(year) + ", save to " + cur_result_xml)
            get_nsfc_data(year, cur_result_xml, autoverify=True)
            last_count += trans2csv(year, cur_result_xml, result_csv, append=True, startnum=(last_count+1))
            year_count += 1
            if(os.path.exists(cur_result_xml)):
                os.remove(cur_result_xml)
        except Exception as e:
            print(str(e))
            print("Get " + str(year) + " fail, retring...")
            continue

    print("Done! Got " + str(last_count) + " items, result was saved to " + result_csv)


    
