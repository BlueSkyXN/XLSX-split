import pandas as pd
import re
import sys
import os

# 读取Excel文件
if len(sys.argv) < 2:
    print("用法: python 客户分类切割.py <输入文件路径.xlsx> [输出目录]")
    print("示例: python 客户分类切割.py 客户对象导出结果.xlsx")
    sys.exit(1)

input_file = sys.argv[1]
if not os.path.exists(input_file):
    print(f"错误: 文件不存在 - {input_file}")
    sys.exit(1)

# 获取输出目录（如果提供了第二个参数）
if len(sys.argv) >= 3:
    output_dir = sys.argv[2]
else:
    output_dir = os.path.dirname(input_file) if os.path.dirname(input_file) else '.'

# 确保输出目录存在
os.makedirs(output_dir, exist_ok=True)

print(f"读取文件: {input_file}")
print(f"输出目录: {output_dir}")

df = pd.read_excel(input_file)
print(f"总共读取到 {len(df)} 行数据")

# 定义规则
rules = {
    "教育类客户": {
        "keywords": ["学校", "大学", "学院", "中学", "小学", "幼儿园", "教育", "实验学校", "研究所", "培训中心", "语言学校", 
                     "技术学院", "职业学校", "高中", "初中", "国际学校", "民办学校", "高等学校", "教育集团", "专科学院", 
                     "专科学校", "大专", "中专", "培训", "继续教育", "成人教育", "职业技术学院", "附属学校", "实验班", 
                     "教育学院", "国际教育"],
        "regex": r".*(小|中|大|专|校)$"
    },
    "公司类客户": {
        "keywords": ["有限", "股份", "集团", "有限公司", "公司", "企业", "实业", "科技", "贸易", "电子", "制造", "工贸", 
                     "工业", "咨询", "财务", "物流", "金融", "资产", "投资", "控股", "商务", "建筑", "设备", "商贸", "合伙企业", 
                     "有限合伙", "事务所", "工作室", "矿业", "煤炭", "钢铁", "能源", "环保", "文化", "传媒", "影视", "广告", 
                     "设计", "会计", "律师", "法律"],
    },
    "医院类客户": {
        "keywords": ["医院", "医疗", "中心", "门诊", "诊所", "保健", "疗养", "急救", "卫生院", "医学院", "疾控", "疾病控制", 
                     "康复", "血站", "体检", "检验", "妇幼保健", "临床", "诊疗", "医务所", "卫生", "健康", "体检", "中医院", 
                     "妇科医院", "专科医院", "肿瘤医院", "眼科", "皮肤科", "精神病院", "检测中心", "研究所", "省立", "协和", "妇幼"],
    },
    "政府党行政机构类客户": {
        "keywords": ["政府", "局", "委员会", "办公室", "部", "处", "厅", "党", "党委", "机关", "事务局", "管理局", "监察", 
                     "公安", "税务", "海关", "国土", "法院", "人民检察院", "行政", "中央", "国家", "地方", "政治", "政法", 
                     "社区", "街道", "办事处", "实名", "中心", "村委会", "监管", "办公室", "联合会", "指挥部", "指挥中心", 
                     "工作组", "督导组", "检查站", "委", "办", "监狱", "警",  "检察院"],
    },
    "差旅酒店类客户": {
        "keywords": ["酒店", "旅馆", "宾馆", "住宿", "客栈", "度假村", "旅社", "旅居", "旅店", "连锁酒店", "商务酒店", 
                     "青年旅舍", "招待所", "民宿", "度假", "休闲", "温泉", "会馆", "山庄", "客房", "国际机场", "旅行社", 
                     "出入境服务", "连锁", "希尔顿", "万豪", "洲际", "会议中心", "展览中心"],
    },
    "电信通讯运营商类客户": {
        "keywords": ["电信", "通信", "移动", "联通", "网络", "互联", "数据", "运营", "通讯", "信息", "广电", "互联网", 
                     "卫星", "天线", "传输", "信号", "服务商", "运营商", "传感", "光纤", "中移", "中电", "中联", "广播", 
                     "电视", "通信", "数据", "物联网", "人工智能", "5G", "云服务", "信息技术"],
    },
    "金融银行类客户": {
        "keywords": [
            "银行", "金融", "投资", "证券", "保险", "信托", "基金", "资产管理", "财务", "支付", "结算", "贷款", 
            "融资", "信贷", "担保", "信用", "风险", "财产", "货币", "外汇", "分行", "支行", "总行", "营业部", 
            "网点", "借贷", "小额贷款", "保险经纪", "信贷公司", "小额贷款公司", "财富管理", "国际银行", 
            "投资银行", "融资租赁", "农商行", "中行", "建行", "工行", "农行", "交行", "招行", "中信", "兴业", 
            "浦发", "光大", "民生", "华夏", "广发", "平安", "浙商", "恒丰", "渤海", "邮储"
        ],
    },
    "中字头或国字头企业": {
        "keywords": ["中华", "国家", "中国", "中央", "中共", "国资", "央企", "国企", "烟", "电网", "电力", "电网", "邮政", "总工会"],
        "regex": r"^(中|国)"
    },
    "交通运输物流仓储类企业": {
        "keywords": ["交通", "运输", "物流", "仓储", "快递", "航运", "航空", "货运", "铁路", "公路", "海运", "货代", 
                     "车队", "集运", "船务", "港务", "运输公司", "物流公司", "仓库", "配送", "机场", "海关", "码头", 
                     "港口", "船舶", "航站", "飞机", "轮船", "车站", "高铁", "轻轨", "地铁", "城际铁路", "物流园", 
                     "保税仓", "邮政"],
    }
}

# 开关：True表示将所有子表写入一个新的Excel文件的不同标签页中
save_in_one_file = True

# 创建一个空的DataFrame来保存不匹配的数据
unmatched = df.copy()

if save_in_one_file:
    output_file = os.path.join(output_dir, '客户名称分类结果.xlsx')
    writer = pd.ExcelWriter(output_file, engine='xlsxwriter')

# 创建子表
for category, rule in rules.items():
    # 初始化筛选条件
    condition = pd.Series(False, index=df.index)
    
    # 处理关键词匹配
    if "keywords" in rule:
        condition |= df['Name'].str.contains('|'.join(rule['keywords']), case=False, na=False)
    
    # 处理正则表达式匹配
    if "regex" in rule:
        condition |= df['Name'].str.match(rule['regex'])
    
    # 根据条件提取数据
    matched = df[condition]
    
    # 将匹配到的从未匹配中删除
    unmatched = unmatched.loc[~condition]
    
    if save_in_one_file:
        # 将匹配到的数据写入同一个Excel文件的不同标签页
        matched.to_excel(writer, sheet_name=category, index=False)
        print(f"分类 '{category}': {len(matched)} 个客户")
    else:
        # 将子表保存为单独的Excel文件
        output_file = os.path.join(output_dir, f'{category}_客户名称.xlsx')
        matched.to_excel(output_file, index=False)
        print(f"分类 '{category}': {len(matched)} 个客户，已保存到: {output_file}")

if save_in_one_file:
    # 将未匹配到的数据写入最后一个标签页
    unmatched.to_excel(writer, sheet_name='未匹配客户名称', index=False)
    writer.close()
    print(f"所有分类结果已保存到: {os.path.join(output_dir, '客户名称分类结果.xlsx')}")
else:
    # 保存未匹配到的客户名称
    output_file = os.path.join(output_dir, '未匹配客户名称.xlsx')
    unmatched.to_excel(output_file, index=False)
    print(f"未匹配客户名称已保存到: {output_file}")

print(f"处理完成！输出目录: {output_dir}")
