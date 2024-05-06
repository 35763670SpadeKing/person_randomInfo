# -*- coding: utf-8 -*-
import os
import string
import datetime
import pandas as pd
from openpyxl import Workbook
# import radar
import random
from faker import Faker


fake = Faker('zh_CN')

class BatchInsert:
    # 职业字典
    position = "医生,老师,飞行员,邮递员,警察,护士,科学家,美术家,艺术家,歌手,办公室职员,经理,老板,助手".split(",")
    # 姓氏字典
    firstName = "赵钱孙李周吴郑王冯陈褚卫蒋沈韩杨朱秦尤许何吕施张孔曹严华金魏陶姜戚谢邹喻柏水窦章云苏潘葛奚范彭郎鲁韦昌马苗凤花方俞任袁柳酆鲍史唐费廉岑薛雷贺倪汤滕殷罗毕" \
                "郝邬安常乐于时傅皮卞齐康伍余元卜顾孟平黄和穆萧尹姚邵湛汪祁毛禹狄米贝明臧计伏成戴谈宋茅庞熊纪舒屈项祝董梁杜阮蓝闵席季麻强贾路娄危江童颜郭梅盛林刁钟徐邱骆高"
    # 女生名字字典
    girl = "秀娟英华慧巧美娜静淑惠珠翠雅芝玉萍红娥玲芬芳燕彩春菊兰凤洁梅琳素云莲真环雪荣" \
           "爱妹霞香月莺媛艳瑞凡佳嘉琼勤珍贞莉桂娣叶璧璐娅琦晶妍茜秋珊莎锦黛青倩婷姣婉娴瑾" \
           "颖露瑶怡婵雁蓓纨仪荷丹蓉眉君琴蕊薇菁梦岚苑婕馨瑗琰韵融园艺咏卿聪澜纯毓悦昭冰爽" \
           "琬茗羽希宁欣飘育滢馥筠柔竹霭凝晓欢霄枫芸菲寒伊亚宜可姬舒影荔枝思丽 "
    # 男生名字字典
    boy = "伟刚勇毅俊峰强军平保东文辉力明永健世广志义兴良海山仁波宁贵福生龙元全国胜学祥才发武新利清飞彬富顺信子杰涛昌成康" \
          "星光天达安岩中茂进林有坚和彪博诚先敬震振壮会思群豪心邦承乐绍功松善厚庆磊民友裕河哲江超浩亮政谦亨奇固之轮翰朗伯" \
          "宏言若鸣朋斌梁栋维启克伦翔旭鹏泽晨辰士以建家致树炎德行时泰盛雄琛钧冠策腾楠榕风航弘 "
    # 道路地址字典
    road = "江南大厦,吴兴一广场,珠海二街,嘉峪关路,高邮湖街,湛山三路,澳门六广场,泰州二路,东海一大厦,天台二路,微山湖街,洞庭湖广场,珠海支街,福州南路,澄海二街,泰州四路,香港中大厦,澳门五路,新湛三街,澳门一路," \
           "正阳关街,宁武关广场,闽江四街,新湛一路,宁国一大厦,王家麦岛,澳门七广场,泰州一路,泰州六街,大尧二路,青大一街,闽江二广场,闽江一大厦,屏东支路,湛山一街,东海西路,徐家麦岛函谷关广场,大尧三路,晓望支街," \
           "秀湛二路,逍遥三大厦,澳门九广场,泰州五街,澄海一路,澳门八街,福州北路,珠海一广场,宁国二路,临淮关大厦,燕儿岛路,紫荆关街,武胜关广场,逍遥一街,秀湛四路,居庸关街,山海关路,鄱阳湖大厦,新湛路,漳州街," \
           "仙游路,花莲街,乐清广场,巢湖街,台南路,吴兴大厦,新田路,福清广场,澄海路,莆田街,海游路,镇江街,石岛广场,宜兴大厦,三明路,仰口街,沛县路,漳浦广场,大麦岛,台湾街,天台路,金湖大厦,高雄广场,海江街," \
           "岳阳路,善化街,荣成路,澳门广场,武昌路,闽江大厦,台北路,龙岩街,咸阳广场,宁德街,龙泉路,丽水街,海川路,彰化大厦,金田路,泰州街,太湖路,江西街,泰兴广场,青大街,金门路,南通大厦,旌德路,汇泉广场," \
           "宁国路,泉州街,如东路,奉化街,鹊山广场,莲岛大厦,华严路,嘉义街,古田路,南平广场,秀湛路,长汀街,湛山路,徐州大厦,丰县广场,汕头街,新竹路,黄海街,安庆路,基隆广场,韶关路,云霄大厦,新安路,仙居街," \
           "屏东广场,晓望街,海门路,珠海街,上杭路,永嘉大厦,漳平路,盐城街,新浦路,新昌街,高田广场,市场三街,金乡东路,市场二大厦,上海支路,李村支广场,惠民南路,市场纬街,长安南路,陵县支街,冠县支广场," \
           "小港一大厦,市场一路,小港二街,清平路,广东广场,新疆路,博平街,港通路,小港沿,福建广场,高唐街,茌平路,港青街,高密路,阳谷广场,平阴路,夏津大厦,邱县路,渤海街,恩县广场,旅顺街,堂邑路,李村街,即墨路," \
           "港华大厦,港环路,馆陶街,普集路,朝阳街,甘肃广场,港夏街,港联路,陵县大厦,上海路,宝山广场,武定路,长清街,长安路,惠民街,武城广场,聊城大厦,海泊路,沧口街,宁波路,胶州广场,莱州路,招远街,冠县路," \
           "六码头,金乡广场,禹城街,临清路,东阿街,吴淞路,大港沿,辽宁路,棣纬二大厦,大港纬一路,贮水山支街,无棣纬一广场,大港纬三街,大港纬五路,大港纬四街,大港纬二路,无棣二大厦,吉林支路,大港四街,普集支路," \
           "无棣三街,黄台支广场,大港三街,无棣一路,贮水山大厦,泰山支路,大港一广场,无棣四路,大连支街,大港二路,锦州支街,德平广场,高苑大厦,长山路,乐陵街,临邑路,嫩江广场,合江路,大连街,博兴路,蒲台大厦," \
           "黄台广场,城阳街,临淄路,安邱街,临朐路,青城广场,商河路,热河大厦,济阳路,承德街,淄川广场,辽北街,阳信路,益都街,松江路,流亭大厦,吉林路,恒台街,包头路,无棣街,铁山广场,锦州街,桓台路,兴安大厦," \
           "邹平路,胶东广场,章丘路,丹东街,华阳路,青海街,泰山广场,周村大厦,四平路,台东西七街,台东东二路,台东东七广场,台东西二路,东五街,云门二路,芙蓉山村,延安二广场,云门一街,台东四路,台东一街,台东二路," \
           "杭州支广场,内蒙古路,台东七大厦,台东六路,广饶支街,台东八广场,台东三街,四平支路,郭口东街,青海支路,沈阳支大厦,菜市二路,菜市一街,北仲三路,瑞云街,滨县广场,庆祥街,万寿路,大成大厦,芙蓉路,历城广场," \
           "大名路,昌平街,平定路,长兴街,浦口广场,诸城大厦,和兴路,德盛街,宁海路,威海广场,东山路,清和街,姜沟路,雒口大厦,松山广场,长春街,昆明路,顺兴街,利津路,阳明广场,人和路,郭口大厦,营口路,昌邑街," \
           "孟庄广场,丰盛街,埕口路,丹阳街,汉口路,洮南大厦,桑梓路,沾化街,山口路,沈阳街,南口广场,振兴街,通化路,福寺大厦,寿光广场,曹县路,昌乐街,道口路,南九水街,台湛广场,东光大厦,驼峰路,太平山,标山路," \
           "云溪广场,太清路".split(",")
    # 邮箱后缀字典
    email_suffix = "@gmail.com,@yahoo.com,@msn.com,@hotmail.com,@live.com,@qq.com,@0355.net,@163.com,@163.net," \
                   "@263.net,@3721.net,@yeah.net,@googlemail.com,@126.com,@sina.com,@sohu.com,@yahoo.com.cn".split(",")


# 生成随机职业
def GetJob():
    job = random.choice(BatchInsert.position)
    return job

# 生成随机性别
def GetGender():
    sex = random.choice("MF")
    return sex

# 获取浙江省所有区县的地区代码，可以替换身份证前6位
def generate_random_zhejiang_area_code():
    zhejiang_area_codes = {
        '杭州市': {
            '上城区': '330102',
            '下城区': '330103',
            '江干区': '330104',
            '拱墅区': '330105',
            '西湖区': '330106',
            '滨江区': '330108',
            '萧山区': '330109',
            '余杭区': '330110',
            '富阳区': '330111',
            '临安区': '330112',
            '桐庐县': '330122',
            '淳安县': '330127',
            '建德市': '330182'
        },
        '宁波市': {
            '海曙区': '330203',
            '江东区': '330204',
            '江北区': '330205',
            '北仑区': '330206',
            '镇海区': '330211',
            '鄞州区': '330212',
            '象山县': '330225',
            '宁海县': '330226',
            '余姚市': '330281',
            '慈溪市': '330282',
            '奉化区': '330213'
        },
        '温州市': {
            '鹿城区': '330302',
            '龙湾区': '330303',
            '瓯海区': '330304',
            '洞头区': '330305',
            '永嘉县': '330324',
            '平阳县': '330326',
            '苍南县': '330327',
            '文成县': '330328',
            '泰顺县': '330329',
            '瑞安市': '330381',
            '乐清市': '330382'
        },
        '嘉兴市': {
            '南湖区': '330402',
            '秀洲区': '330411',
            '嘉善县': '330421',
            '海盐县': '330424',
            '海宁市': '330481',
            '平湖市': '330482',
            '桐乡市': '330483'
        },
        '湖州市': {
            '吴兴区': '330502',
            '南浔区': '330503',
            '德清县': '330521',
            '长兴县': '330522',
            '安吉县': '330523'
        },
        '绍兴市': {
            '越城区': '330602',
            '柯桥区': '330603',
            '上虞区': '330604',
            '新昌县': '330624',
            '诸暨市': '330681',
            '嵊州市': '330683'
        },
        '金华市': {
            '婺城区': '330702',
            '金东区': '330703',
            '武义县': '330723',
            '浦江县': '330726',
            '磐安县': '330727',
            '兰溪市': '330781',
            '义乌市': '330782',
            '东阳市': '330783',
            '永康市': '330784'
        },
        '衢州市': {
            '柯城区': '330802',
            '衢江区': '330803',
            '常山县': '330822',
            '开化县': '330824',
            '龙游县': '330825',
            '江山市': '330881'
        },
        '舟山市': {
            '定海区': '330902',
            '普陀区': '330903',
            '岱山县': '330921',
            '嵊泗县': '330922'
        },
        '台州市': {
            '椒江区': '331002',
            '黄岩区': '331003',
            '路桥区': '331004',
            '玉环市': '331021',
            '三门县': '331022',
            '天台县': '331023',
            '仙居县': '331024',
            '温岭市': '331081',
            '临海市': '331082'
        },
        '丽水市': {
            '莲都区': '331102',
            '青田县': '331121',
            '缙云县': '331122',
            '遂昌县': '331123',
            '松阳县': '331124',
            '云和县': '331125',
            '庆元县': '331126',
            '景宁畲族自治县': '331127',
            '龙泉市': '331181'
        }
    }

    # 设置温州市的概率为80%，其他地方的概率为20%
    if random.random() < 0.8:
        city_name = '温州市'
    else:
        city_name = random.choice(list(zhejiang_area_codes.keys()))
    district = random.choice(list(zhejiang_area_codes[city_name].keys()))
    area_code = zhejiang_area_codes[city_name][district]

    return area_code


# 生成随机姓名和性别, 值为1生成姓名，值为2生成姓名和性别,值为3返回姓名，性别，身份证
def GetName(option):
    full_name = random.choice(BatchInsert.firstName)
    name_sex = GetGender()  # 随机性别
    is_middle = int(random.choice("0123"))
    if name_sex == 'F' and is_middle > 0:
        last_name = random.choice(BatchInsert.boy) + random.choice(BatchInsert.boy)
    elif name_sex == 'F' and is_middle <= 0:
        last_name = random.choice(BatchInsert.boy)
    elif name_sex == 'M' and is_middle > 0:
        last_name = random.choice(BatchInsert.girl) + random.choice(BatchInsert.girl)
    else:
        last_name = random.choice(BatchInsert.girl)
    full_name = full_name + last_name
    id_number = fake.ssn(min_age = 20, max_age = 70, gender = name_sex)
    if option == 1:
        return full_name
    elif option == 2:
        return full_name, name_sex
    elif option == 3:
        return full_name, id_number.replace(id_number[:6], generate_random_zhejiang_area_code())


# 生成一个指定长度的随机字符串
def random_str(str_length):
    str_list = [random.choice(string.digits + string.ascii_letters) for i in range(str_length)]
    result_str = ''.join(str_list)
    return result_str

# 生成一个随机身份证号码,faker函数的出生时间不可自定义。
def random_id_card():
    id_number = fake.ssn(min_age = 20, max_age= 70)
    return id_number


# 生成随机日期
def get_random_time():
    start_time = datetime.time(8, 30)  # 设置开始时间为 8:30 AM
    end_time = datetime.time(17, 0)  # 设置结束时间为 5:00 PM
    current_date = datetime.datetime.now().date()  # 获取当天日期
    start_datetime = datetime.datetime.combine(current_date, start_time)
    end_datetime = datetime.datetime.combine(current_date, end_time)

    random_time = start_datetime + datetime.timedelta(seconds=random.randint(0, int((end_datetime - start_datetime).total_seconds())))

    return random_time.strftime("%Y-%m-%d %H:%M:%S")



# 生成一个指定长度的随机数字
def random_int(int_length):
    str_list = [random.choice(string.digits) for i in range(int_length)]
    result_int = ''.join(str_list)
    return result_int

# 生成一个62开头19位数字的银行卡号。
def get_bank_card():
    num = "62" + "".join(random.choice("0123456789") for _ in range(17))
    return num

# 生成一个指定区间的随机数字
def getNum(start, end):
    return random.randint(start, end)

def getMarried():
    marital_status = ['已婚', '未婚']
    return random.choice(marital_status)

# 汇总所有信息
def Get_all():
    Res = [  # random_str(8),
        # random_int(9),
        GetName(1),  # 姓名
        random_id_card(),  # 身份证
        GetJob(),   # 职业
        getMarried(), # 婚姻状况
        generate_random_zhejiang_area_code(),   # 邮编
        fake.address(),  # 牌照
        fake.license_plate(),
        fake.phone_number(),  # 手机号
        str(get_bank_card()),  # 银行卡
        str(fake.credit_card_number(card_type=None)),  # 信用卡
        # fake.coordinate(center=None, radius=0.001),  # 坐标
        fake.ascii_free_email(),  # 免费邮箱
        fake.ipv4(network=False, address_class=None, private=None),  # ip地址
        get_random_time()]
    return Res


if __name__ == '__main__':
    # 写入文件的地址
    file_path = 'D:/Software/PyCharm/projects/RandomInfo/data.xlsx'
    if os.path.isfile(file_path):
        # 如果文件已经存在，则打开Excel文件
        df_existing = pd.read_excel(file_path)
    else:
        df_empty = pd.DataFrame()
        df_empty.to_excel(file_path, index=False)
        print("空的D:/Software/PyCharm/projects/RandomInfo/data.xlsx文件已创建")
    # 写入数据
    df_new = pd.DataFrame(
        columns=["Name", "ID Card", "Occupation", "Married","Zip Code", "Address", "License Plate", "Mobile Number", "Bank Card",
                 "Credit Card", "Email", "IP Address", "Time"])

    for i in range(1000):
        # 生成新数据
        data = Get_all()
        # 将新数据添加到DataFrame中
        df_new.loc[i] = data
    # 将新数据追加到已有数据后面
    df_updated = df_existing._append(df_new, ignore_index=True)

    # 写入更新后的数据到Excel文件
    df_updated.to_excel('D:/Software/PyCharm/projects/RandomInfo/data.xlsx', index=False)