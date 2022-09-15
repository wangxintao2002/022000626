import os
import re

provinceList = [
    "北京", "天津", "上海", "重庆", "内蒙古", "广西", "西藏", "宁夏", "新疆", "河北", "山西", "辽宁", "吉林",
    "黑龙江", "江苏", "浙江", "安徽", "福建", "江西", "山东", "河南", "湖北", "湖南", "广东", "河南", "海南",
    "四川", "贵州", "云南", "陕西", "甘肃", "青海"]


def get_all_filename(path):
    return os.listdir(path)


# def process_province(province, d, str):
#     if province.find('，') == -1:
#         for name in provinceList:
#             if name in province:
#                 d[name] = int(d[str])
#                 return
#                 #print(name,d[name])
#     else:
#         if province.find("均") == 0 or province.find("在") == 0:
#             for name in provinceList:
#                 if name in province:
#                     d[name] = int(d[str])
#                     return
#         piece = province.split('，')
#         for one in piece:
#             for name in provinceList:
#                 if name in one:
#                     if (name not in d) and one[one.rfind(name)+len(name)].isdigit():
#                         d[name] = int(one[one.rfind(name)+len(name):one.rfind("例")])
#                     break

alldata = dict()

def process_reg(province,d , str):
    if province.find('，') == -1:
        for name in provinceList:
            if name in province:
                d[name] = int(d[str])
                return
                #print(name,d[name])
    else:
        if province.find("均") == 0 or province.find("在") == 0:
            for name in provinceList:
                if name in province:
                    d[name] = int(d[str])
                    return
        for name in provinceList:
            pattern = re.compile(r"%s[0-9]*例"%name)
            data = re.search(pattern,province)
            if data:
                data = data.group()
                d[name] = int(data[data.find(name)+len(name):data.find('例')])


def process_new_diagnosis(filenames, basePath):
    for filename in filenames:
        date = filename[0:10]
        print(date)
        d = dict()
        alldata[date] = dict()
        file = open(basePath + "\\" + filename, 'r', encoding='utf-8')
        content = file.read()
        pattern = re.compile(r'本土病例.*?）')
        result = re.search(pattern, content)
        if result:
            # 求出本土新增总病例
            num = result.group()
            d['diagnosis'] = int(num[4:num.find("例（")])
            #print(d['diagnosis'])
            # 求出各省份新增病例
            provinces = result.group()
            provinces = provinces[provinces.index('（') + 1:provinces.index("）")]
            #print(provinces)
            process_reg(provinces, d,'diagnosis')

        else:
            d['diagnosis'] = 0
        alldata[date]['diagnosis'] = d

def process_new_asymptomatic(filenames,basePath):
    for filename in filenames:
        date = filename[0:10]
        #print(date)
        d = {'Date': date}
        file = open(basePath + "\\" + filename, 'r', encoding='utf-8')
        content = file.read()
        pattern = re.compile(r'本土\d.*）(?=。\n当日解除|；当日解除|；当日无)')
        result = re.search(pattern,content)
        if result:
            num = result.group()
            d['asymptomatic'] = int(num[2:num.find("例")])
            provinces = result.group()
            provinces = provinces[provinces.find("（")+1:provinces.find("）")]
            #print(provinces)
            process_reg(provinces,d, 'asymptomatic')
            #print(d)

        else:
            d['asymptomatic'] = 0
        alldata[date]['asymptomatic'] = d

if "__main__" == __name__:
    basePath = "D:\\疫情防控数据"
    filenames = get_all_filename(basePath)
    process_new_diagnosis(filenames, basePath)
    process_new_asymptomatic(filenames,basePath)
    print(alldata)

