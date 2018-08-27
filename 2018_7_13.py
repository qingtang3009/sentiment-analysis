# -*- coding: utf-8 -*-

import pandas
from openpyxl import load_workbook
import re
from snownlp import SnowNLP
import jieba
import os
import numpy

pattern0 = u'漂亮|好看|打眼|可爱|颜色|款式|搭配|身材|很酷|配色|帅气|时尚|霸气|橙色|超酷|黑色|眼前一亮|搭|酷|着装|白色|红色|黄色|帅呆了|' \
           u'颜值|精良|大气|美观|结实|轻巧|质感|设计|手感|做工|漂亮|好看|打眼|可爱|颜色|款式|搭配|身材|很酷|配色|帅气|时尚|霸气|橙色|' \
           u'超酷|黑色|体积小|眼前一亮|搭|酷|着装|腰带|白色|红色|黄色|上装|帅呆了|体积小|颜值|精良|大气|美观|结实|轻巧|小巧玲珑|质感|' \
           u'设计|外观|小巧|手感|做工|重量|大小|体积|体积小|小巧玲珑|体积|大小|外观|小巧|重量|轻|轻便'
pattern1 = u'漂亮|好看|打眼|可爱|颜色|款式|搭配|身材|很酷|配色|帅气|时尚|霸气|橙色|超酷|黑色|眼前一亮|搭|酷|着装|白色|红色|黄色|帅呆了|' \
           u'颜值|精良|大气|美观|结实|轻巧|质感|设计|手感|做工|' \
           u'漂亮|好看|打眼|可爱|颜色|款式|搭配|身材|很酷|配色|帅气|时尚|霸气|橙色|超酷|黑色|体积小|眼前一亮|搭|酷|着装|腰带|白色|红色|' \
           u'黄色|上装|帅呆了|体积小|颜值|精良|大气|美观|结实|轻巧|小巧玲珑|质感|设计|外观|小巧|手感|做工|重量|大小|体积'
pattern2 = u'体积小|小巧玲珑|体积|大小|外观|小巧'
pattern3 = u'重量|轻|轻便'
pattern4 = u'说明书|不会|说明图解|清单|介绍|速度|转速|动力|强劲|冲击|档|扭力|扭矩|转速|功率|马力|挡位|调速|力矩|力量|冲击力|力道|力气|' \
           u'冲击|动力|强劲|冲击|档|操作|拧紧|电池|充电|充电器|电量|续航|无线|电池容量|锂电池|usb|20v|20伏|12伏|两用|多功能|开关|' \
           u'电机|噪音|灯|电线|电源|噪声|插头|碳刷|电能|电源线|马达|散热|零件|照明灯|小修|停转|停机|故障|功能|手柄|自锁|调速'
pattern5 = u'说明书|不会|说明图解|清单|介绍'
pattern6 = u'速度|转速'
pattern7 = u'动力|强劲|冲击|档|扭力|扭矩|转速|功率|马力|挡位|调速|力矩|力量|冲击力|力道|力气|冲击|动力|强劲|冲击|档|操作|拧紧'
pattern8 = u'电池|充电|充电器|电量|续航|无线|电池容量|锂电池|usb|20v|20伏|12伏'
pattern9 = u'两用|多功能'
pattern10 = u'开关|电机|噪音|灯|电线|电源|噪声|插头|碳刷|电能|电源线|马达|散热|零件|照明灯|小修|停转|停机|故障|功能|手柄|自锁|调速'
pattern11 = u'螺丝批头|扳手|螺丝起子|夹头|套筒|钻头|螺丝刀头|螺丝刀|钻夹|劈头|磁铁|锈|螺丝头|吸铁石|套装|配件|刀具|锯条|锯片|砂纸|刀片|' \
            u'螺丝|螺丝刀|螺丝批|扳手|螺丝起子|夹头|套筒|电钻|钻头|批头'
pattern12 = u'灰尘|尘|噪音|声音|震动|晃动|偏心|离心|同心度|同轴度|定心|摆头|扭动|震动|抖动'
pattern13 = u'灰尘|尘'
pattern14 = u'噪音|声音'
pattern15 = u'震动|晃动|偏心|离心|同心度|同轴度|定心|摆头|扭动|震动|抖动'
pattern16 = u'包装|包装盒|实物|外包装|电钻盒|包装|说明书|包装盒|实物|外包装|瑕疵|防伪|规格|标签'
pattern17 = u'客服|差评|退货|退款|返现|服务|服务态度|服务周到|售后服务'
pattern18 = u'物流|快递|发货|货|送货|到货|售后|收货|配送|顺丰|开箱|货品|货物|包裹|运送|运输|圆通|订单|急用|商品质量|采购|换货|运费|' \
            u'送货上门|拆封|包装箱|售前|验货|邮费|上门|拆包|拆箱|邮政|物流配送|保证质量|转运|装卸|速递|开包|收件|返款|货运|订货|出库|' \
            u'取件|发回去|退还|货单|ems|EMS'
pattern19 = u'品牌|正品|威克士|大牌子|防伪|Bosch|博世|大品牌|博士'
pattern20 = u'价格|便宜|划算|特价|活动|打折'
pattern21 = u'20v|20伏|12伏'

pattern_list = [pattern0, pattern1, pattern2, pattern3, pattern4, pattern5, pattern6, pattern7, pattern8, pattern9,
                pattern10, pattern11, pattern12, pattern13, pattern14, pattern15, pattern16, pattern17, pattern18,
                pattern19, pattern20, pattern21]

attribute_list = ['外观', '外观（设计）', '外观（体积）', '外观（重量）', '产品', '产品（指引）', '产品（转速）', '产品（力量）',
                  '产品（电池）', '产品（多功能）', '产品（产品其他）', '附件', '使用', '使用（灰尘）', '使用（声音）', '使用（晃动）',
                  '包装', '客服', '物流', '品牌', '价格', '可能涉及其他产品']

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

def xlsx2list(excel_file_name):
    """
     读取xlsx格式文件

    :param excel_file_name: 爬取的评论数据
    :return:
    [
    [用户1评论]，
    [用户2评论，附加评论]
    ]
    ，
    [时间0，时间1，时间2，......]
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
    读取CSV格式文件

    :param csv_file_name:
    :return:
    [
    [用户1评论]，
    [用户2评论，附加评论]
    ]
    ，
    [时间0，时间1，时间2，......]
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


def trans(word_in_comment):
    """
    繁简转换

    :param word_in_comment:
    :return: 繁转简，去除&sheel类似的字符
    """
    new_word = SnowNLP(word_in_comment).han
    new_word = re.sub(r'&[a-z]*', '', new_word)
    new_word = re.sub(r'\ufffd', '', new_word)
    return new_word


def cut(sentence):
    """
    分词，去除停用词

    :param sentence: 评论
    :return: 词列表
    """
    jieba.load_userdict('./jieba_dict/user_dictionary.txt')
    with open('./stopwords.txt', 'r', encoding='UTF-8') as input_file:
        stopwords = input_file.readlines()
    words = []
    for word in jieba.__lcut(sentence):
        if word not in stopwords:
            words.append(word)
    return words


def split_short(long_comment):
    """
    将长评论按pattern切分成短评论

    :param long_comment: 单条长评论
    :return: [str1,str2,str3]
    """
    pattern = u'，+|。+|！+|？+|\\n|\.+|,+|!+|\?+|;+|；+|、+|但|\s+'
    p = re.compile(pattern)
    comments = p.split(long_comment)
    # 去除空字符串
    short_comment_list = []
    for comment in comments:
        if len(list(comment)) > 0:
            short_comment_list.append(comment)
    return short_comment_list


def pro_or_diy(comments_list):
    """
    判断是业余人士购买（diy）或专业人士购买（professional）

    :param comments_list:从xlsx或者csv读取后的大列表
    :return:每个长评论对应的类别
    """
    pattern1 = r'帮[\u4E00-\u9FA5]+买|给[\u4E00-\u9FA5]+买|购置|采购|买给|购买[\u4E00-\u9FA5]+'

    pattern1_1 = r'同事|公司|师傅|单位|客户|车间|水电|安装工|仓管|机修|工程师|职业|职业维修工|木工|干维修的|专业的|' \
                 r'跑工地|搞维修|维修工作|员工|安防工程|做装修的|裱装|装裱'

    pattern1_2 = r'哥哥|老婆|老公|爸|姐|夫|伙伴|自用|居家|家庭|日常|家里|家居|家用|DIY|diy|在家|家人|兴趣小组|修剪下冬青|院子里的树|' \
                 r'树|装个监控|装个画架|照片墙|灯|安装监控摄像头|宜家|搬家|初用者|不用求人|雇人|没有用过电钻|女生|拜托别人|小白|工地'

    pattern2 = r'安装|修理|制作|工程|打孔|机修|家具|装修|柜子|书柜|书架|墙|砖|改造|瓷砖|水泥|混凝土'

    pattern2_1 = r'工程|修理工|职业|工友|从事|技师|师傅|工厂|公司|修理店|技术人员|机修|同事|公司|师傅|单位|客户|工人|' \
                 r'车间|水电|安装工|仓管|机修|工程师|电工|职业|职业维修工|木工|干维修的|跑工地|搞维修|维修工作|员工|安防工程|做装修的'

    pattern2_2 = r'居家|家庭|日常|家里|家居|家用|DIY|diy|在家|家人|兴趣小组|修剪下冬青|院子里的树|树|装个监控|装个画架|照片墙|灯|自己动手|自动搞定|' \
                 r'安装监控摄像头|宜家|搬家|初用者|不用求人|雇人|没有用过电钻|女生|拜托别人|~|女汉子|女生|女汉子|老公|～|女|娘|姐|丈夫|~'

    pattern3 = '起停|无极调速|切割深度|锉磨|铲切|抛光|用过好几个|悬空握持|三角夹头|工序|跳动|比国产|比杂牌'
    pattern3_3 = '习惯好评|吃灰|偶尔|学会|不会用|居家|不太会用|试试|买来玩'

    # 因为使用场景、购买场景和后续的关键词可能分别出现在初次评论和附加评论，为了逻辑方便，先把两个句子合并
    string_list = []
    for i in range(len(comments_list)):
        sent = ''
        for one_string in comments_list[i]:
            sent += one_string
        string_list.append(sent)
    result = []

    for sent in string_list:
        if re.compile(pattern1).search(sent) and re.compile(pattern1_1).search(sent):
            result.append(1)
        elif re.compile(pattern2).search(sent) and re.compile(pattern2_1).search(sent):
            result.append(2)
        elif re.compile(pattern3).search(sent):
            result.append(3)
        elif re.compile(pattern1_1).search(sent) or re.compile(pattern2_1).search(sent):
            result.append(3)
        elif re.compile(pattern1).search(sent) and re.compile(pattern1_2).search(sent):
            result.append(-1)
        elif re.compile(pattern2).search(sent) and re.compile(pattern2_2).search(sent):
            result.append(-2)
        elif re.compile(pattern3_3).search(sent):
            result.append(-3)
        elif re.compile(pattern1_2).search(sent) or re.compile(pattern2_2).search(sent):
            result.append(-3)
        else:
            result.append(0)
    return result


def score(comment):
    """
    给comment打分，出现正向词+1，出现负向词-1，出现反转词乘以-1，最后汇总

    :param comment:短评论
    :return: 每条短评论对应的分数
    """
    positive_words = []
    negative_words = []
    with open(r'D:\Pycharm\PycharmProjects\Class\qinggancihui\negative.txt', 'r', encoding='utf-8') as neg_input:
        for word in neg_input.readlines():
            negative_words.append(word.strip())
    # print(negative_words)
    with open(r'D:\Pycharm\PycharmProjects\Class\qinggancihui\positive.txt', 'r', encoding='utf-8') as pos_input:
        for word in pos_input.readlines():
            positive_words.append(word.strip())
    # print(positive_words)

    positive_extra = ['价格便宜', '不贵', '价格公道', '价格合理', '太棒了', '比较满意',
                      '便携', '完全一致', '省时省力', '强烈推荐', '一分价钱一分货', '强劲有力',
                      '价美物廉', '非常适合', '服务周到', '非常感谢',
                      '不贵', '方便使用', '价格合理', '很正', '劲道', '轻巧方便', '妥妥',
                      '质量上乘', '够快', '经济实用', '不愧为', '特快', '美观大方', '大大提高', '小爽', '惊艳', '运气',
                      '好极了', '非常简单', '认真负责', '一如既往地好', '价格公道', '特别感谢', '灰常好', '不震手', '一应俱全',
                      '不错呀', '太牛', '别看', '很强', '强', '对得起', '不错']
    negative_extra = ['不值', '小贵', '美中不足', '破损', '很差', '差', '太差', '鸡肋', '别买', '态度恶劣', '不合理', '太不给力', '偏软', '不够', '差劲', '坑']

    negative_words.extend(negative_extra)
    positive_words.extend(positive_extra)

    positive_words.remove('默认')
    positive_words.remove('需要')
    positive_words.remove('知道')

    negative_words.remove('活动')
    negative_words.remove('没有')
    negative_words.remove('很强')
    negative_words.remove('强')
    negative_words.remove('买好')
    negative_words.remove('杂牌')
    negative_words.remove('细')
    negative_words.remove('轻易')
    negative_words.remove('酷')

    transition = ['不', '没有', '不如', '不是', '无', '没']

    pos_pattern = u'噪音[\u4E00-\u9FA5]*低|声音不大|体积[\u4E00-\u9FA5]*小|体积不大|声音[\u4E00-\u9FA5]*小|声音[\u4E00-\u9FA5]*轻|做工细|马力[\u4E00-\u9FA5]*大|扭矩[\u4E00-\u9FA5]*大|噪音小|噪声小|重量轻|操作简单'
    neg_pattern = u'进灰尘|有灰尘|噪声[\u4E00-\u9FA5]*大|噪音[\u4E00-\u9FA5]*大|包装简单|灰尘[\u4E00-\u9FA5]*大|声音[\u4E00-\u9FA5]*大'
    pos_p = re.compile(pos_pattern)
    neg_p = re.compile(neg_pattern)

    if pos_p.search(comment):
        print(comment)
        return 1
    elif neg_p.search(comment):
        print(comment)
        return -1

    words = cut(comment)

    neg_count = 0
    pos_count = 0
    z = 1
    for word in words:
        if (word in negative_words) is True:
            neg_count = neg_count + 1
        elif (word in positive_words) is True:
            pos_count = pos_count + 1
        elif (word in transition) is True:
            z *= -1
    if (neg_count - pos_count) > 0:
        return z * -1
    elif (neg_count - pos_count) < 0:
        return z * 1
    else:
        return 0


def judge(score_list, sentences):
    """
    判断每个总的短句子中的每个小类的正负向，正向大于福相为1，负向大于正向为-1，正负向均没有为0。
    特例是如果碰到一个长评论切分的短评论包含对一个类的正负向情感，此时赋值0.1

    :param score_list:
    :param sentences:
    :return:
    """
    pos_and_neg_ = []
    if len(score_list) == 0:
        return 0
    flag_pos = 0
    flag_neg = 0
    for i in score_list:
        if i > 0:
            flag_pos += 1
        elif i < 0:
            flag_neg += 1
    if flag_pos and flag_neg:
        sentence_ = [cut(sent) for sent in sentences]
        pos_and_neg_.append((score_list, sentence_))
        return 0.1
    elif flag_pos:
        return 1
    elif flag_neg:
        return -1
    else:
        return 0


def make_sheet(comment_list, pattern_list):
    """
    制作表格，传入所有长评论和pattern_list，将长评论切分成短评论，给每个短评论打分，加上上面的judge，返回制成表格形式

    :param comment_list: 所有评论
    :param pattern_list: 所有pattern
    :return: 分数表格
    """
    data_sheet = numpy.zeros((len(comment_list), len(pattern_list)))

    short_comment = []
    for i in range(len(comment_list)):
        short_comment.append(split_short(comment_list[i]))
    # print(short_comment)
    for i in range(len(pattern_list)):
        p = re.compile(pattern_list[i])
        for j in range(len(short_comment)):
            one_type_comments = []
            score_list = []
            for one_sent in short_comment[j]:
                if p.search(one_sent):
                    one_type_comments.append(one_sent)
            for comment in one_type_comments:
                score_list.append(score(comment))
            data_sheet[j, i] = judge(score_list=score_list, sentences=one_type_comments)
    print(data_sheet)
    return data_sheet


def distinguish(comment_list):
    """
    将评论中的初始评论和附加评论切分开

    :param comment_list: 从xlsx和csv读取的评论
    :return: 初始评论，附加评论
    """
    first_comment = []
    append_comment = []
    for i in range(len(comment_list)):
        first_comment.append(comment_list[i][0])
        if len(comment_list[i]) > 1:
            append_comment.append(comment_list[i][1])
        else:
            append_comment.append("")
    print(first_comment, append_comment)
    return first_comment, append_comment


def data_tocsv(input_file_name, comment_list, date_list, pattern_list, pattern_name, shangjia):
    """
    制表，分为初始评论数据表和附加评论数据表

    :param input_file_name: 输入文件名
    :param comment_list: 评论列表
    :param date_list: 时间列表
    :param pattern_list: pattern列表
    :param pattern_name: pattern名
    :param shangjia: 商家ID
    :return: 两个表格
    """
    global file2name

    # 评论总数
    comment_sum = len(comment_list)

    first_comment, append_comment = distinguish(comment_list)

    # 评论类别统计
    data_sheet = make_sheet(comment_list=first_comment, pattern_list=pattern_list)
    data_sheet2 = make_sheet(comment_list=append_comment, pattern_list=pattern_list)

    # 统计一下每个类别的个数

    total = numpy.sum(data_sheet[:, 1:4], axis=1) + numpy.sum(data_sheet[:, 5:12], axis=1) + numpy.sum(
        data_sheet[:, 13:], axis=1)
    # 将横向量变为竖向量
    total = total.reshape(-1, )

    data2 = pandas.DataFrame(data_sheet, columns=pattern_name)
    # 在统计矩阵的datafram 中把评论的数据加进去
    data2['正面/负面'] = pandas.Series(total)
    data2['PRO/DIY'] = pandas.Series(pro_or_diy(comment_list))
    data2['时间'] = pandas.Series(date_list)
    data2['商家'] = pandas.Series([shangjia for i in range(comment_sum)])
    data2['名称'] = pandas.Series([file2name[input_file_name] for i in range(comment_sum)])
    data2["评论"] = pandas.Series(first_comment)

    total1 = numpy.sum(data_sheet2[:, 1:4], axis=1) + numpy.sum(data_sheet2[:, 5:12], axis=1) + numpy.sum(
        data_sheet2[:, 13:], axis=1)

    total1 = total1.reshape(-1, )
    data3 = pandas.DataFrame(data_sheet2, columns=pattern_name)
    # 在统计矩阵的datafram 中把评论的数据加进去
    data3['正面/负面'] = pandas.Series(total1)
    data3['PRO/DIY'] = pandas.Series(pro_or_diy(comment_list))
    data3['时间'] = pandas.Series(date_list)
    data3['商家'] = pandas.Series([shangjia for i in range(comment_sum)])
    data3['名称'] = pandas.Series([file2name[input_file_name] for i in range(comment_sum)])
    data3["附加评论"] = pandas.Series(append_comment)

    return data2, data3


if __name__ == '__main__':
    PATH = './data_6_30/'
    list_dir = os.listdir(PATH)
    all_pd2 = []
    all_pd3 = []
    for file_name in list_dir[:7]:
        print(file_name)
        file_data, date = xlsx2list(PATH + file_name)
        data2, data3 = data_tocsv(input_file_name=file_name, comment_list=file_data, date_list=date,
                                  pattern_list=pattern_list, pattern_name=attribute_list, shangjia='JD')
        all_pd2.append(data2)
        all_pd3.append(data3)
    for file_name in list_dir[7:]:
        print(file_name)
        file_data, date = csv2list(PATH + file_name)
        data2, data3 = data_tocsv(input_file_name=file_name, comment_list=file_data, date_list=date,
                                  pattern_list=pattern_list, pattern_name=attribute_list, shangjia='TMALL')
        all_pd2.append(data2)
        all_pd3.append(data3)

    data_frame_concat = pandas.concat(all_pd2, axis=0, ignore_index=True)
    data_frame_concat.to_excel('7_113_all.xlsx', index=False, encoding='gb18030')

    data_frame_concat = pandas.concat(all_pd3, axis=0, ignore_index=True)
    data_frame_concat.to_excel('7_13_append.xlsx', index=False, encoding='gb18030')
