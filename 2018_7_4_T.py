# -*- coding: utf-8 -*-

import re
import pandas
import numpy
import os
from openpyxl import load_workbook
from snownlp import SnowNLP

file2name = {
             '6_29_jd_2349393.xlsx': 'WX128',
             '6_29_jd_3518165.xlsx': 'WX679',
             '6_29_jd_3940357.xlsx': 'WX382',
             '6_29_jd_3949250.xlsx': 'WX317 ',
             '6_29_jd_4576540.xlsx': 'WX550',
             '6_29_jd_560885.xlsx': 'GSB600RE',
             '6_29_jd_5853353.xlsx': 'Bosch GO',
             '6_29_TMall_1040076538_38087781618.csv': 'GSR120',
             '6_29_TMall_1040076538_44694284453.csv': 'TSB5500',
             '6_29_TMall_1734094296_38569773211.csv': 'WX252',
             '6_29_TMall_1734094296_38652440512.csv': 'WX382',
             '6_29_TMall_1734094296_45160197777.csv': 'WX679',
             '6_29_TMall_1734094296_521063319115.csv': 'WX128',
             '6_29_TMall_1734094296_541600800605.csv': 'WX550',
             '6_29_TMall_2493613286_528759674980.csv': 'GSR120',
             '6_29_TMall_2587124438_521406153256.csv': 'Bosch GO',
             '6_29_TMall_707360547_38143181564.csv': 'GSR120',
             '6_29_TMall_707360547_44400807101.csv': 'TSB5500'
            }
pattern = u'漂亮|好看|打眼|可爱|颜色|款式|搭配|身材|很酷|配色|帅气|时尚|霸气|橙色|超酷|黑色|眼前一亮|搭|酷|着装|白色|红色|黄色|帅呆了|颜值|精良|大气|美观|结实|轻巧|质感|设计|手感|做工|漂亮|好看|打眼|可爱|颜色|款式|搭配|身材|很酷|配色|帅气|时尚|霸气|橙色|超酷|黑色|体积小|眼前一亮|搭|酷|着装|腰带|白色|红色|黄色|上装|帅呆了|体积小|颜值|精良|大气|美观|结实|轻巧|小巧玲珑|质感|设计|外观|小巧|手感|做工|重量|大小|体积|体积小|小巧玲珑|体积|大小|外观|小巧|重量|轻|轻便'
pattern0 = u'漂亮|好看|打眼|可爱|颜色|款式|搭配|身材|很酷|配色|帅气|时尚|霸气|橙色|超酷|黑色|眼前一亮|搭|酷|着装|白色|红色|黄色|帅呆了|颜值|精良|大气|美观|结实|轻巧|质感|设计|手感|做工|' \
           u'漂亮|好看|打眼|可爱|颜色|款式|搭配|身材|很酷|配色|帅气|时尚|霸气|橙色|超酷|黑色|体积小|眼前一亮|搭|酷|着装|腰带|白色|红色|黄色|上装|帅呆了|体积小|颜值|精良|大气|美观|结实|轻巧|小巧玲珑|质感|设计|外观|小巧|手感|做工|重量|大小|体积'
pattern1 = u'体积小|小巧玲珑|体积|大小|外观|小巧'
pattern2 = u'重量|轻|轻便'
pattern_add = u'说明书|不会|说明图解|清单|介绍|速度|转速|动力|强劲|冲击|档|扭力|扭矩|转速|功率|马力|挡位|调速|力矩|力量|冲击力|力道|力气|冲击|动力|强劲|冲击|档|操作|拧紧|电池|充电|充电器|电量|续航|无线|电池容量|锂电池|usb|20v|20伏|12伏|两用|多功能|开关|电机|噪音|灯|电线|电源|噪声|插头|碳刷|电能|电源线|马达|散热|零件|照明灯|小修|停转|停机|故障|功能|手柄|自锁|调速'
pattern3 = u'说明书|不会|说明图解|清单|介绍'
pattern4 = u'速度|转速'
pattern5 = u'动力|强劲|冲击|档|扭力|扭矩|转速|功率|马力|挡位|调速|力矩|力量|冲击力|力道|力气|冲击|' \
           u'动力|强劲|冲击|档|操作|拧紧'
pattern6 = u'电池|充电|充电器|电量|续航|无线|电池容量|锂电池|usb|20v|20伏|12伏'
pattern7 = u'两用|多功能'
pattern8 = u'开关|电机|噪音|灯|电线|电源|噪声|插头|碳刷|电能|电源线|马达|散热|零件|照明灯|小修|停转|停机|故障|功能|手柄|自锁|调速'
pattern9 = u'螺丝批头|扳手|螺丝起子|夹头|套筒|钻头|螺丝刀头|螺丝刀|钻夹|劈头|磁铁|锈|螺丝头|吸铁石|套装|配件|刀具|锯条|锯片|砂纸|刀片|' \
           u'螺丝|螺丝刀|螺丝批|扳手|螺丝起子|夹头|套筒|电钻|钻头|批头'
pattern222 = u'灰尘|尘|噪音|声音|震动|晃动|偏心|离心|同心度|同轴度|定心|摆头|扭动|震动|抖动'
pattern10 = u'灰尘|尘'
pattern11 = u'噪音|声音'
pattern12 = u'震动|晃动|偏心|离心|同心度|同轴度|定心|摆头|扭动|震动|抖动'
pattern13 = u'包装|包装盒|实物|外包装|电钻盒|' \
            u'包装|说明书|包装盒|实物|外包装|瑕疵|防伪|规格|标签'
pattern14 = u'客服|差评|退货|退款|返现|服务|服务态度|服务周到|售后服务'
pattern15 = u'物流|快递|发货|货|送货|到货|售后|收货|配送|顺丰|开箱|货品|货物|包裹|运送|运输|圆通|订单|急用|商品质量|采购|换货|运费|' \
            u'送货上门|拆封|包装箱|售前|验货|邮费|上门|拆包|拆箱|邮政|物流配送|保证质量|转运|装卸|速递|开包|收件|返款|货运|订货|出库|' \
            u'取件|发回去|退还|货单|ems|EMS'
pattern16 = u'品牌|正品|威克士|大牌子|防伪|Bosch|博世|大品牌|博士'
pattern17 = u'价格|便宜|划算|特价|活动|打折'
pattern18 = u'20v|20伏|12伏'

pattern_list = [pattern, pattern0, pattern1, pattern2, pattern_add, pattern3, pattern4, pattern5, pattern6, pattern7, pattern8, pattern9,
                pattern222, pattern10, pattern11, pattern12, pattern13, pattern14, pattern15, pattern16, pattern17, pattern18]

attribute_list = ['外观', '外观（设计）', '外观（体积）', '外观（重量）', '产品', '产品（指引）', '产品（转速）', '产品（力量）', '产品（电池）',
                  '产品（多功能）', '产品（产品其他）', '附件', '使用', '使用（灰尘）', '使用（声音）', '使用（晃动）', '包装',
                  '客服', '物流', '品牌', '价格', '可能涉及其他产品']


def trans(sent):
    """
    繁体简体转换
    和参数相关的词转化为<param> 这里会把部分商品型号也转化为<param>
    删除&hell 这些无效符合
    转化数字为<num>
    """
    new_sent = SnowNLP(sent).han
    new_sent = re.sub(r'&[a-z]*', '', new_sent)
    new_sent = re.sub(r'\ufffd', '', new_sent)
    return new_sent


def xlsx2list(excel_file_name):
    """
    xlsx 文件转换为list
    为进行有无附加评论和无效评论的筛选

    xlsx 转化为list之后
    对list进行无效评论删除
    最后转化为
    [
    [用户1评论]，
    [用户2评论，附加评论]
    ]
    """
    wb = load_workbook(excel_file_name)
    sheet = wb.active
    # 获得当前正在显示的sheet, 也可以用wb.get_active_sheet()
    return_list = []
    for row in sheet.rows:
        one_row = []
        for cell in row:
            one_row.append(cell.value)
        return_list.append(one_row)
    print(return_list)

    comment_list = []
    pattern1 = r'此用户没有填写评论|此用户未填写评价内容|此用户未及时评价'
    date_list = []
    p = re.compile(pattern1)
    for i in range(1, len(return_list)):
        one_line = []
        flag = p.search(str(return_list[i][6]))
        if not flag:
            one_line.append(trans(return_list[i][6]))
        # 判断有无 追加评论
        if return_list[i][7] is not None:
            flag = p.search(str(return_list[i][7]))
            if not flag:
                one_line.append(trans(return_list[i][7]))
        if len(one_line) > 0:
            date_list.append(return_list[i][1])
            comment_list.append(one_line)
    print(comment_list, date_list)
    return comment_list, date_list


def csv2list(csv_file_name):
    """
    读取CSV 转化为：包含评论的list 和 时间list
    [
    [用户1评论]，
    [用户2评论，附加评论]
    ]
    转化过程删除无效评论
    """
    pattern0 = r'此用户没有填写评论|此用户未填写评价内容|此用户未及时评价'
    p = re.compile(pattern0)
    input_info = pandas.read_csv(csv_file_name, encoding='gbk')
    a = list(input_info.loc[:, ['appendComment', 'rateContent']].values)
    b = list(input_info.loc[:, ['rateDate']].values)
    date_list = []
    all_list = []
    pattern = re.compile(r"content': '.+?'")
    for i in range(len(a)):
        one_list = []
        if not p.search(str(a[i][1])):
            one_list.append(trans(str(a[i][1])))
        if not a[i][0]:
            if not p.search(str(a[i][0])):
                match = pattern.search(a[i][0])
                print(type(match))
                one_list.append(trans(match.group(0)[11:-1]))
        if len(one_list) > 0:
            date_list.append(str(b[i])[2:-2])
            all_list.append(one_list)
    return all_list, date_list


def good_or_bad(comment_list):
    '''
    with open(r'D:\Pycharm\PycharmProjects\Class\qinggancihui\positive.txt', 'r', encoding='utf-8') as i1:
        p1 = i1.readlines()
    with open(r'D:\Pycharm\PycharmProjects\Class\qinggancihui\negative.txt', 'r', encoding='utf-8') as i2:
        p2 = i2.readlines()
    '''
    p1 = r'好|不错|可以|还行|棒'
    p2 = r'差|不行|垃圾|太差'
    pattern_good = re.compile(p1)
    pattern_bad = re.compile(p2)
    list_for_count = []
    # dictionary = {}
    string_list = []
    for i in range(len(comment_list)):
        sent = ''
        for one_string in comment_list[i]:
            sent += one_string
        string_list.append(sent)
    for comment in string_list:
        sigh = len(re.findall(pattern_good, comment)) - len(re.findall(pattern_bad, comment))
        if sigh > 0:
            list_for_count.append(1)
        elif sigh == 0:
            list_for_count.append(0)
        else:
            list_for_count.append(-1)

    '''
        if sigh > 0:
            dictionary[comment] = 1
        elif sigh == 0:
            dictionary[comment] = 0
        else:
            dictionary[comment] = -1
    '''
    # print(list_for_count)
    return list_for_count


def pro_or_diy(comment_list):
    pattern1 = r'帮[\u4E00-\u9FA5]+买|给[\u4E00-\u9FA5]+买|购置|采购|买给|购买[\u4E00-\u9FA5]+'
    pattern1_1 = r'同事|公司|师傅|单位|客户|工人|车间|水电|安装工|仓管|机修|工程师|工作|电工|专业|职业|职业维修工|木工|干维修的|' \
                 r'跑工地|搞维修|维修工作|员工|安防工程|做装修的'
    pattern1_2 = r'哥哥|老婆|老公|爸|姐|夫|伙伴|自用|居家|家庭|日常|家里|家居|家用|DIY|diy|在家|家人|兴趣小组|修剪下冬青|院子里的树|' \
                 r'树|装个监控|装个画架|照片墙|灯|安装监控摄像头|宜家|搬家|初用者|不用求人|雇人|没有用过电钻|女生|拜托别人'
    pattern2 = r'安装|修理|制作|工程|打孔|机修|家具|装修|柜子|书柜|书架|墙|砖|改造|瓷砖|水泥|混凝土'
    pattern2_1 = r'工程|修理工|职业|高空作业|工友|从事|技师|师傅|厂|工厂|公司|修理店|技术人员|机修|干活|同事|公司|师傅|单位|客户|工人|' \
                 r'车间|水电|安装工|仓管|机修|工程师|工作|电工|专业|职业|职业维修工|木工|干维修的|跑工地|搞维修|维修工作|员工|安防工程|做装修的'
    pattern2_2 = r'居家|家庭|日常|家里|家居|家用|DIY|diy|在家|家人|兴趣小组|修剪下冬青|院子里的树|树|装个监控|装个画架|照片墙|灯|' \
                 r'安装监控摄像头|宜家|搬家|初用者|不用求人|雇人|没有用过电钻|女生|拜托别人'
    pattern3 = '同心度|同轴度|起停|无极调速|切割深度|锉磨|铲切|抛光|用过好几个|悬空握持|三角夹头|工序|跳动'
    pattern3_3 = '习惯好评|吃灰|偶尔|不多|学会|试了下|不会'

    # 因为 使用场景、购买场景 和后续的关键词可能分别出现在 初次评论和附加评论
    # 为了逻辑方便，先把两个句子合并
    string_list = []
    for i in range(len(comment_list)):
        sent = ''
        for one_string in comment_list[i]:
            sent += one_string
        string_list.append(sent)
    result = []
    # 判断
    for sent in string_list:
        if re.compile(pattern1).search(sent) and re.compile(pattern1_1).search(sent):
            result.append(1)
        elif re.compile(pattern2).search(sent) and re.compile(pattern2_1).search(sent):
            result.append(2)
        elif re.compile(pattern3).search(sent):
            result.append(3)
        elif re.compile(pattern1_1).search(sent) or re.compile(pattern2_1).search(sent):
            result.append(4)
        elif re.compile(pattern1).search(sent) and re.compile(pattern1_2).search(sent):
            result.append(-1)
        elif re.compile(pattern2).search(sent) and re.compile(pattern2_2).search(sent):
            result.append(-2)
        elif re.compile(pattern3_3).search(sent):
            result.append(-3)
        elif re.compile(pattern1_2).search(sent) or re.compile(pattern2_2).search(sent):
            result.append(-4)
        else:
            result.append(0)
    return result


def data_tocsv(input_file_name, comment_list, date_list, pattern_list, pattern_name, shangjia):
    """
    将传入comment_list 转化为pd.dataframe
    input_file_name是传入的文件名‘xxxxx.csv  xxxxx.xlsx’
    comment list 是评论的列表 可以是
    [
     [评论1]
     [评论2,评论2附加评论]
    ]
    对于comment list 如果格式不一样 下面for的格式改一下

    patten_list=[pattern0,pattern1,pattern2,pattern3,pattern4,pattern5,pattern6,pattern7,pattern8]

    patten_name=['Packaging','Performance','Appearance',\
    'Product','Product (usage)','Product (battery)','Accessories','Customer service','Logistics']

    """
    global file2name

    # 生成统计用的numpy
    data_sheet = numpy.zeros((len(comment_list), len(pattern_list)), dtype=int)

    # 评论类别统计
    for i in range(len(pattern_list)):
        p = re.compile(pattern_list[i])
        for j in range(len(comment_list)):  # 这里for的格式改一下
            for one_sent in comment_list[j]:
                if p.search(one_sent):
                    data_sheet[j, i] = 1
                    break
    comment_sum = len(comment_list)
    data2 = pandas.DataFrame(data_sheet, columns=pattern_name)
    # 在统计矩阵的datafram 中把评论的数据加进去
    data2['情感'] = pandas.Series(good_or_bad(comment_list))
    data2['PRO/DIY'] = pandas.Series(pro_or_diy(comment_list))
    data2['时间'] = pandas.Series(date_list)
    data2['商家'] = pandas.Series([shangjia for i in range(comment_sum)])
    data2['名称'] = pandas.Series([file2name[input_file_name] for i in range(comment_sum)])
    data2["评论"] = pandas.Series(comment_list)
    return data2


if __name__ == '__main__':
    PATH = './data_6_30/'
    list_dir = os.listdir(PATH)
    all_pd = []
    for file_name in list_dir[:7]:
        print(file_name)
        file_data, date = xlsx2list(PATH+file_name)
        data_pd = data_tocsv(input_file_name=file_name, comment_list=file_data, date_list=date,
                             pattern_list=pattern_list, pattern_name=attribute_list, shangjia='JD')
        all_pd.append(data_pd)
    for file_name in list_dir[7:]:
        print(file_name)
        file_data, date = csv2list(PATH+file_name)
        data_pd = data_tocsv(input_file_name=file_name, comment_list=file_data,
                             date_list=date, pattern_list=pattern_list, pattern_name=attribute_list,
                             shangjia='TMALL')
        all_pd.append(data_pd)

    data_frame_concat = pandas.concat(all_pd, axis=0, ignore_index=True)
    data_frame_concat.to_excel('7_10_all_T.xlsx', index=False, encoding='gb18030')
