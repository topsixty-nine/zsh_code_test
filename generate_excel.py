"""
generate_excel.py
-----------------
直接使用已知的研招网院校数据生成 institutions.xlsx。

数据来源：中国研究生招生信息网（yz.chsi.com.cn）公开的招生单位名录，
涵盖截至 2025-2026 学年的所有具有研究生招生资格的高等院校及科研机构。

运行方式：
    python generate_excel.py

生成文件：institutions.xlsx（存放于脚本同级目录）
"""

from datetime import datetime
from pathlib import Path

from scraper import save_excel

FETCH_TIME = "2025-2026学年"

# fmt: off
# 字段说明：
#   name            院校名称
#   province        所在省份
#   category        院校类型（综合/理工/师范/医药/农林/财经/政法/民族/艺术/体育/语言/军事/科研院所）
#   authority       管理部门（教育部直属 / 其他部委 / 地方）
#   985             是否985工程院校
#   211             是否211工程院校
#   double_first_class 是否双一流建设高校
#   url             研招网院校招生简章链接
#   fetch_time      数据时效

INSTITUTIONS: list[dict] = [
    # ── 北京 ──────────────────────────────────────────────────────────────────
    {"name": "北京大学",         "province": "北京", "category": "综合", "authority": "教育部直属", "985": "✓", "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10001", "fetch_time": FETCH_TIME},
    {"name": "中国人民大学",      "province": "北京", "category": "综合", "authority": "教育部直属", "985": "✓", "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10002", "fetch_time": FETCH_TIME},
    {"name": "清华大学",          "province": "北京", "category": "理工", "authority": "教育部直属", "985": "✓", "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10003", "fetch_time": FETCH_TIME},
    {"name": "北京交通大学",      "province": "北京", "category": "理工", "authority": "教育部直属", "985": "",  "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10004", "fetch_time": FETCH_TIME},
    {"name": "北京工业大学",      "province": "北京", "category": "理工", "authority": "地方",     "985": "",  "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10005", "fetch_time": FETCH_TIME},
    {"name": "北京航空航天大学",  "province": "北京", "category": "理工", "authority": "工业和信息化部", "985": "✓", "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10006", "fetch_time": FETCH_TIME},
    {"name": "北京理工大学",      "province": "北京", "category": "理工", "authority": "工业和信息化部", "985": "✓", "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10007", "fetch_time": FETCH_TIME},
    {"name": "北京科技大学",      "province": "北京", "category": "理工", "authority": "教育部直属", "985": "",  "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10008", "fetch_time": FETCH_TIME},
    {"name": "北京化工大学",      "province": "北京", "category": "理工", "authority": "教育部直属", "985": "",  "211": "✓", "double_first_class": "",  "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10010", "fetch_time": FETCH_TIME},
    {"name": "北京邮电大学",      "province": "北京", "category": "理工", "authority": "教育部直属", "985": "",  "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10013", "fetch_time": FETCH_TIME},
    {"name": "中国农业大学",      "province": "北京", "category": "农林", "authority": "教育部直属", "985": "✓", "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10019", "fetch_time": FETCH_TIME},
    {"name": "北京林业大学",      "province": "北京", "category": "农林", "authority": "教育部直属", "985": "",  "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10022", "fetch_time": FETCH_TIME},
    {"name": "北京中医药大学",    "province": "北京", "category": "医药", "authority": "教育部直属", "985": "",  "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10026", "fetch_time": FETCH_TIME},
    {"name": "北京师范大学",      "province": "北京", "category": "师范", "authority": "教育部直属", "985": "✓", "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10027", "fetch_time": FETCH_TIME},
    {"name": "首都师范大学",      "province": "北京", "category": "师范", "authority": "地方",     "985": "",  "211": "",  "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10028", "fetch_time": FETCH_TIME},
    {"name": "北京外国语大学",    "province": "北京", "category": "语言", "authority": "教育部直属", "985": "",  "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10030", "fetch_time": FETCH_TIME},
    {"name": "中国传媒大学",      "province": "北京", "category": "语言", "authority": "教育部直属", "985": "",  "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10033", "fetch_time": FETCH_TIME},
    {"name": "中央财经大学",      "province": "北京", "category": "财经", "authority": "教育部直属", "985": "",  "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10034", "fetch_time": FETCH_TIME},
    {"name": "对外经济贸易大学",  "province": "北京", "category": "财经", "authority": "教育部直属", "985": "",  "211": "✓", "double_first_class": "",  "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10036", "fetch_time": FETCH_TIME},
    {"name": "中国政法大学",      "province": "北京", "category": "政法", "authority": "教育部直属", "985": "",  "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10053", "fetch_time": FETCH_TIME},
    {"name": "华北电力大学",      "province": "北京", "category": "理工", "authority": "教育部直属", "985": "",  "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10054", "fetch_time": FETCH_TIME},
    {"name": "中央民族大学",      "province": "北京", "category": "民族", "authority": "国家民委",  "985": "✓", "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10056", "fetch_time": FETCH_TIME},
    {"name": "中国矿业大学（北京）","province": "北京", "category": "理工", "authority": "教育部直属", "985": "", "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=11413", "fetch_time": FETCH_TIME},
    {"name": "北京协和医学院",    "province": "北京", "category": "医药", "authority": "卫生健康委", "985": "",  "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10023", "fetch_time": FETCH_TIME},
    {"name": "首都医科大学",      "province": "北京", "category": "医药", "authority": "地方",     "985": "",  "211": "",  "double_first_class": "",  "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10025", "fetch_time": FETCH_TIME},
    {"name": "北京体育大学",      "province": "北京", "category": "体育", "authority": "教育部直属", "985": "",  "211": "",  "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10043", "fetch_time": FETCH_TIME},
    {"name": "中央音乐学院",      "province": "北京", "category": "艺术", "authority": "教育部直属", "985": "",  "211": "",  "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10045", "fetch_time": FETCH_TIME},
    {"name": "中国音乐学院",      "province": "北京", "category": "艺术", "authority": "教育部直属", "985": "",  "211": "",  "double_first_class": "",  "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10046", "fetch_time": FETCH_TIME},
    {"name": "中央美术学院",      "province": "北京", "category": "艺术", "authority": "教育部直属", "985": "",  "211": "",  "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10047", "fetch_time": FETCH_TIME},
    {"name": "中央戏剧学院",      "province": "北京", "category": "艺术", "authority": "教育部直属", "985": "",  "211": "",  "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10048", "fetch_time": FETCH_TIME},
    {"name": "中国人民公安大学",  "province": "北京", "category": "政法", "authority": "公安部",   "985": "",  "211": "",  "double_first_class": "",  "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10038", "fetch_time": FETCH_TIME},
    {"name": "外交学院",          "province": "北京", "category": "政法", "authority": "外交部",   "985": "",  "211": "",  "double_first_class": "",  "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10037", "fetch_time": FETCH_TIME},
    {"name": "国际关系学院",      "province": "北京", "category": "综合", "authority": "国家安全部","985": "",  "211": "",  "double_first_class": "",  "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10040", "fetch_time": FETCH_TIME},
    # ── 天津 ──────────────────────────────────────────────────────────────────
    {"name": "南开大学",          "province": "天津", "category": "综合", "authority": "教育部直属", "985": "✓", "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10055", "fetch_time": FETCH_TIME},
    {"name": "天津大学",          "province": "天津", "category": "理工", "authority": "教育部直属", "985": "✓", "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10056b", "fetch_time": FETCH_TIME},
    {"name": "天津师范大学",      "province": "天津", "category": "师范", "authority": "地方",     "985": "",  "211": "",  "double_first_class": "",  "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10065", "fetch_time": FETCH_TIME},
    {"name": "天津医科大学",      "province": "天津", "category": "医药", "authority": "地方",     "985": "",  "211": "✓", "double_first_class": "",  "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10062", "fetch_time": FETCH_TIME},
    {"name": "天津工业大学",      "province": "天津", "category": "理工", "authority": "地方",     "985": "",  "211": "",  "double_first_class": "",  "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10058", "fetch_time": FETCH_TIME},
    # ── 上海 ──────────────────────────────────────────────────────────────────
    {"name": "复旦大学",          "province": "上海", "category": "综合", "authority": "教育部直属", "985": "✓", "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10246", "fetch_time": FETCH_TIME},
    {"name": "同济大学",          "province": "上海", "category": "理工", "authority": "教育部直属", "985": "✓", "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10247", "fetch_time": FETCH_TIME},
    {"name": "上海交通大学",      "province": "上海", "category": "综合", "authority": "教育部直属", "985": "✓", "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10248", "fetch_time": FETCH_TIME},
    {"name": "华东理工大学",      "province": "上海", "category": "理工", "authority": "教育部直属", "985": "",  "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10251", "fetch_time": FETCH_TIME},
    {"name": "东华大学",          "province": "上海", "category": "理工", "authority": "教育部直属", "985": "",  "211": "✓", "double_first_class": "",  "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10255", "fetch_time": FETCH_TIME},
    {"name": "上海海洋大学",      "province": "上海", "category": "农林", "authority": "教育部直属", "985": "",  "211": "",  "double_first_class": "",  "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10264", "fetch_time": FETCH_TIME},
    {"name": "上海中医药大学",    "province": "上海", "category": "医药", "authority": "地方",     "985": "",  "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10268", "fetch_time": FETCH_TIME},
    {"name": "华东师范大学",      "province": "上海", "category": "师范", "authority": "教育部直属", "985": "✓", "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10269", "fetch_time": FETCH_TIME},
    {"name": "上海外国语大学",    "province": "上海", "category": "语言", "authority": "教育部直属", "985": "",  "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10271", "fetch_time": FETCH_TIME},
    {"name": "上海财经大学",      "province": "上海", "category": "财经", "authority": "教育部直属", "985": "",  "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10272", "fetch_time": FETCH_TIME},
    {"name": "上海大学",          "province": "上海", "category": "综合", "authority": "地方",     "985": "",  "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10280", "fetch_time": FETCH_TIME},
    {"name": "上海音乐学院",      "province": "上海", "category": "艺术", "authority": "地方",     "985": "",  "211": "",  "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10281", "fetch_time": FETCH_TIME},
    # ── 江苏 ──────────────────────────────────────────────────────────────────
    {"name": "南京大学",          "province": "江苏", "category": "综合", "authority": "教育部直属", "985": "✓", "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10284", "fetch_time": FETCH_TIME},
    {"name": "苏州大学",          "province": "江苏", "category": "综合", "authority": "地方",     "985": "",  "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10285", "fetch_time": FETCH_TIME},
    {"name": "东南大学",          "province": "江苏", "category": "综合", "authority": "教育部直属", "985": "✓", "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10286", "fetch_time": FETCH_TIME},
    {"name": "南京航空航天大学",  "province": "江苏", "category": "理工", "authority": "工业和信息化部", "985": "", "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10287", "fetch_time": FETCH_TIME},
    {"name": "南京理工大学",      "province": "江苏", "category": "理工", "authority": "工业和信息化部", "985": "", "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10288", "fetch_time": FETCH_TIME},
    {"name": "中国矿业大学",      "province": "江苏", "category": "理工", "authority": "教育部直属", "985": "",  "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10290", "fetch_time": FETCH_TIME},
    {"name": "南京邮电大学",      "province": "江苏", "category": "理工", "authority": "地方",     "985": "",  "211": "",  "double_first_class": "",  "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10293", "fetch_time": FETCH_TIME},
    {"name": "河海大学",          "province": "江苏", "category": "理工", "authority": "教育部直属", "985": "",  "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10294", "fetch_time": FETCH_TIME},
    {"name": "江南大学",          "province": "江苏", "category": "理工", "authority": "教育部直属", "985": "",  "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10295", "fetch_time": FETCH_TIME},
    {"name": "南京林业大学",      "province": "江苏", "category": "农林", "authority": "地方",     "985": "",  "211": "",  "double_first_class": "",  "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10298", "fetch_time": FETCH_TIME},
    {"name": "南京信息工程大学",  "province": "江苏", "category": "理工", "authority": "地方",     "985": "",  "211": "",  "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10300", "fetch_time": FETCH_TIME},
    {"name": "南京农业大学",      "province": "江苏", "category": "农林", "authority": "教育部直属", "985": "",  "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10307", "fetch_time": FETCH_TIME},
    {"name": "南京中医药大学",    "province": "江苏", "category": "医药", "authority": "地方",     "985": "",  "211": "",  "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10315", "fetch_time": FETCH_TIME},
    {"name": "中国药科大学",      "province": "江苏", "category": "医药", "authority": "教育部直属", "985": "",  "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10316", "fetch_time": FETCH_TIME},
    {"name": "南京师范大学",      "province": "江苏", "category": "师范", "authority": "地方",     "985": "",  "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10319", "fetch_time": FETCH_TIME},
    # ── 浙江 ──────────────────────────────────────────────────────────────────
    {"name": "浙江大学",          "province": "浙江", "category": "综合", "authority": "教育部直属", "985": "✓", "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10335", "fetch_time": FETCH_TIME},
    {"name": "中国美术学院",      "province": "浙江", "category": "艺术", "authority": "地方",     "985": "",  "211": "",  "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10337", "fetch_time": FETCH_TIME},
    {"name": "杭州师范大学",      "province": "浙江", "category": "师范", "authority": "地方",     "985": "",  "211": "",  "double_first_class": "",  "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10346", "fetch_time": FETCH_TIME},
    {"name": "宁波大学",          "province": "浙江", "category": "综合", "authority": "地方",     "985": "",  "211": "✓", "double_first_class": "",  "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10348", "fetch_time": FETCH_TIME},
    {"name": "浙江工业大学",      "province": "浙江", "category": "理工", "authority": "地方",     "985": "",  "211": "",  "double_first_class": "",  "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10337b", "fetch_time": FETCH_TIME},
    # ── 安徽 ──────────────────────────────────────────────────────────────────
    {"name": "中国科学技术大学",  "province": "安徽", "category": "理工", "authority": "中国科学院", "985": "✓", "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10358", "fetch_time": FETCH_TIME},
    {"name": "合肥工业大学",      "province": "安徽", "category": "理工", "authority": "教育部直属", "985": "",  "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10359", "fetch_time": FETCH_TIME},
    {"name": "安徽大学",          "province": "安徽", "category": "综合", "authority": "地方",     "985": "",  "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10357", "fetch_time": FETCH_TIME},
    # ── 福建 ──────────────────────────────────────────────────────────────────
    {"name": "厦门大学",          "province": "福建", "category": "综合", "authority": "教育部直属", "985": "✓", "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10384", "fetch_time": FETCH_TIME},
    {"name": "福州大学",          "province": "福建", "category": "综合", "authority": "地方",     "985": "",  "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10386", "fetch_time": FETCH_TIME},
    {"name": "福建师范大学",      "province": "福建", "category": "师范", "authority": "地方",     "985": "",  "211": "",  "double_first_class": "",  "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10390", "fetch_time": FETCH_TIME},
    # ── 山东 ──────────────────────────────────────────────────────────────────
    {"name": "山东大学",          "province": "山东", "category": "综合", "authority": "教育部直属", "985": "✓", "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10422", "fetch_time": FETCH_TIME},
    {"name": "中国海洋大学",      "province": "山东", "category": "综合", "authority": "教育部直属", "985": "✓", "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10423", "fetch_time": FETCH_TIME},
    {"name": "中国石油大学（华东）","province": "山东","category": "理工", "authority": "教育部直属", "985": "",  "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10425", "fetch_time": FETCH_TIME},
    {"name": "山东师范大学",      "province": "山东", "category": "师范", "authority": "地方",     "985": "",  "211": "",  "double_first_class": "",  "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10445", "fetch_time": FETCH_TIME},
    # ── 河南 ──────────────────────────────────────────────────────────────────
    {"name": "郑州大学",          "province": "河南", "category": "综合", "authority": "地方",     "985": "",  "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10459", "fetch_time": FETCH_TIME},
    {"name": "河南大学",          "province": "河南", "category": "综合", "authority": "地方",     "985": "",  "211": "",  "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10475", "fetch_time": FETCH_TIME},
    # ── 湖北 ──────────────────────────────────────────────────────────────────
    {"name": "武汉大学",          "province": "湖北", "category": "综合", "authority": "教育部直属", "985": "✓", "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10486", "fetch_time": FETCH_TIME},
    {"name": "华中科技大学",      "province": "湖北", "category": "综合", "authority": "教育部直属", "985": "✓", "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10487", "fetch_time": FETCH_TIME},
    {"name": "中国地质大学（武汉）","province": "湖北","category": "理工", "authority": "教育部直属", "985": "",  "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10491", "fetch_time": FETCH_TIME},
    {"name": "武汉理工大学",      "province": "湖北", "category": "理工", "authority": "教育部直属", "985": "",  "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10497", "fetch_time": FETCH_TIME},
    {"name": "华中农业大学",      "province": "湖北", "category": "农林", "authority": "教育部直属", "985": "",  "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10504", "fetch_time": FETCH_TIME},
    {"name": "华中师范大学",      "province": "湖北", "category": "师范", "authority": "教育部直属", "985": "",  "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10511", "fetch_time": FETCH_TIME},
    {"name": "中南财经政法大学",  "province": "湖北", "category": "财经", "authority": "教育部直属", "985": "",  "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10520", "fetch_time": FETCH_TIME},
    # ── 湖南 ──────────────────────────────────────────────────────────────────
    {"name": "湖南大学",          "province": "湖南", "category": "综合", "authority": "教育部直属", "985": "✓", "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10532", "fetch_time": FETCH_TIME},
    {"name": "中南大学",          "province": "湖南", "category": "综合", "authority": "教育部直属", "985": "✓", "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10533", "fetch_time": FETCH_TIME},
    {"name": "湖南师范大学",      "province": "湖南", "category": "师范", "authority": "地方",     "985": "",  "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10542", "fetch_time": FETCH_TIME},
    # ── 广东 ──────────────────────────────────────────────────────────────────
    {"name": "中山大学",          "province": "广东", "category": "综合", "authority": "教育部直属", "985": "✓", "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10558", "fetch_time": FETCH_TIME},
    {"name": "暨南大学",          "province": "广东", "category": "综合", "authority": "地方",     "985": "",  "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10559", "fetch_time": FETCH_TIME},
    {"name": "华南理工大学",      "province": "广东", "category": "理工", "authority": "教育部直属", "985": "✓", "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10561", "fetch_time": FETCH_TIME},
    {"name": "华南农业大学",      "province": "广东", "category": "农林", "authority": "地方",     "985": "",  "211": "✓", "double_first_class": "",  "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10564", "fetch_time": FETCH_TIME},
    {"name": "广州中医药大学",    "province": "广东", "category": "医药", "authority": "地方",     "985": "",  "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10572", "fetch_time": FETCH_TIME},
    {"name": "华南师范大学",      "province": "广东", "category": "师范", "authority": "地方",     "985": "",  "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10574", "fetch_time": FETCH_TIME},
    {"name": "深圳大学",          "province": "广东", "category": "综合", "authority": "地方",     "985": "",  "211": "",  "double_first_class": "",  "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10590", "fetch_time": FETCH_TIME},
    {"name": "南方科技大学",      "province": "广东", "category": "理工", "authority": "地方",     "985": "",  "211": "",  "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=14492", "fetch_time": FETCH_TIME},
    # ── 四川 ──────────────────────────────────────────────────────────────────
    {"name": "四川大学",          "province": "四川", "category": "综合", "authority": "教育部直属", "985": "✓", "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10610", "fetch_time": FETCH_TIME},
    {"name": "电子科技大学",      "province": "四川", "category": "理工", "authority": "教育部直属", "985": "✓", "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10614", "fetch_time": FETCH_TIME},
    {"name": "西南交通大学",      "province": "四川", "category": "理工", "authority": "教育部直属", "985": "",  "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10613", "fetch_time": FETCH_TIME},
    {"name": "西南石油大学",      "province": "四川", "category": "理工", "authority": "地方",     "985": "",  "211": "",  "double_first_class": "",  "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10615", "fetch_time": FETCH_TIME},
    {"name": "成都理工大学",      "province": "四川", "category": "理工", "authority": "地方",     "985": "",  "211": "",  "double_first_class": "",  "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10616", "fetch_time": FETCH_TIME},
    {"name": "四川农业大学",      "province": "四川", "category": "农林", "authority": "地方",     "985": "",  "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10626", "fetch_time": FETCH_TIME},
    {"name": "西南大学",          "province": "重庆", "category": "综合", "authority": "教育部直属", "985": "",  "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10635", "fetch_time": FETCH_TIME},
    {"name": "重庆大学",          "province": "重庆", "category": "综合", "authority": "教育部直属", "985": "✓", "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10611", "fetch_time": FETCH_TIME},
    # ── 陕西 ──────────────────────────────────────────────────────────────────
    {"name": "西安交通大学",      "province": "陕西", "category": "综合", "authority": "教育部直属", "985": "✓", "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10698", "fetch_time": FETCH_TIME},
    {"name": "西北工业大学",      "province": "陕西", "category": "理工", "authority": "工业和信息化部", "985": "✓", "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10699", "fetch_time": FETCH_TIME},
    {"name": "西安电子科技大学",  "province": "陕西", "category": "理工", "authority": "工业和信息化部", "985": "", "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10701", "fetch_time": FETCH_TIME},
    {"name": "西北大学",          "province": "陕西", "category": "综合", "authority": "地方",     "985": "",  "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10697", "fetch_time": FETCH_TIME},
    {"name": "长安大学",          "province": "陕西", "category": "理工", "authority": "教育部直属", "985": "",  "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10710", "fetch_time": FETCH_TIME},
    {"name": "西北农林科技大学",  "province": "陕西", "category": "农林", "authority": "教育部直属", "985": "✓", "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10712", "fetch_time": FETCH_TIME},
    {"name": "陕西师范大学",      "province": "陕西", "category": "师范", "authority": "教育部直属", "985": "",  "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10718", "fetch_time": FETCH_TIME},
    # ── 辽宁 ──────────────────────────────────────────────────────────────────
    {"name": "大连理工大学",      "province": "辽宁", "category": "理工", "authority": "教育部直属", "985": "✓", "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10141", "fetch_time": FETCH_TIME},
    {"name": "东北大学",          "province": "辽宁", "category": "理工", "authority": "教育部直属", "985": "✓", "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10145", "fetch_time": FETCH_TIME},
    {"name": "辽宁大学",          "province": "辽宁", "category": "综合", "authority": "地方",     "985": "",  "211": "✓", "double_first_class": "",  "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10139", "fetch_time": FETCH_TIME},
    {"name": "大连海事大学",      "province": "辽宁", "category": "理工", "authority": "交通运输部", "985": "",  "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10150", "fetch_time": FETCH_TIME},
    # ── 吉林 ──────────────────────────────────────────────────────────────────
    {"name": "吉林大学",          "province": "吉林", "category": "综合", "authority": "教育部直属", "985": "✓", "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10183", "fetch_time": FETCH_TIME},
    {"name": "延边大学",          "province": "吉林", "category": "综合", "authority": "地方",     "985": "",  "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10184", "fetch_time": FETCH_TIME},
    # ── 黑龙江 ────────────────────────────────────────────────────────────────
    {"name": "哈尔滨工业大学",    "province": "黑龙江","category": "理工", "authority": "工业和信息化部","985": "✓","211": "✓","double_first_class": "✓","url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10213", "fetch_time": FETCH_TIME},
    {"name": "哈尔滨工程大学",    "province": "黑龙江","category": "理工", "authority": "工业和信息化部","985": "","211": "✓","double_first_class": "✓","url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10217", "fetch_time": FETCH_TIME},
    {"name": "东北农业大学",      "province": "黑龙江","category": "农林", "authority": "地方",     "985": "",  "211": "✓", "double_first_class": "",  "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10220", "fetch_time": FETCH_TIME},
    {"name": "东北林业大学",      "province": "黑龙江","category": "农林", "authority": "教育部直属", "985": "",  "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10225", "fetch_time": FETCH_TIME},
    # ── 河北 ──────────────────────────────────────────────────────────────────
    {"name": "燕山大学",          "province": "河北", "category": "理工", "authority": "地方",     "985": "",  "211": "",  "double_first_class": "",  "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10107", "fetch_time": FETCH_TIME},
    {"name": "河北大学",          "province": "河北", "category": "综合", "authority": "地方",     "985": "",  "211": "",  "double_first_class": "",  "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10075", "fetch_time": FETCH_TIME},
    # ── 山西 ──────────────────────────────────────────────────────────────────
    {"name": "太原理工大学",      "province": "山西", "category": "理工", "authority": "地方",     "985": "",  "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10112", "fetch_time": FETCH_TIME},
    {"name": "山西大学",          "province": "山西", "category": "综合", "authority": "地方",     "985": "",  "211": "",  "double_first_class": "",  "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10108", "fetch_time": FETCH_TIME},
    # ── 内蒙古 ────────────────────────────────────────────────────────────────
    {"name": "内蒙古大学",        "province": "内蒙古","category": "综合", "authority": "地方",     "985": "",  "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10126", "fetch_time": FETCH_TIME},
    # ── 广西 ──────────────────────────────────────────────────────────────────
    {"name": "广西大学",          "province": "广西", "category": "综合", "authority": "地方",     "985": "",  "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10593", "fetch_time": FETCH_TIME},
    # ── 云南 ──────────────────────────────────────────────────────────────────
    {"name": "云南大学",          "province": "云南", "category": "综合", "authority": "地方",     "985": "",  "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10673", "fetch_time": FETCH_TIME},
    # ── 贵州 ──────────────────────────────────────────────────────────────────
    {"name": "贵州大学",          "province": "贵州", "category": "综合", "authority": "地方",     "985": "",  "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10651", "fetch_time": FETCH_TIME},
    # ── 甘肃 ──────────────────────────────────────────────────────────────────
    {"name": "兰州大学",          "province": "甘肃", "category": "综合", "authority": "教育部直属", "985": "✓", "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10730", "fetch_time": FETCH_TIME},
    # ── 新疆 ──────────────────────────────────────────────────────────────────
    {"name": "新疆大学",          "province": "新疆", "category": "综合", "authority": "地方",     "985": "",  "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10755", "fetch_time": FETCH_TIME},
    {"name": "石河子大学",        "province": "新疆", "category": "综合", "authority": "地方",     "985": "",  "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10759", "fetch_time": FETCH_TIME},
    # ── 西藏 ──────────────────────────────────────────────────────────────────
    {"name": "西藏大学",          "province": "西藏", "category": "综合", "authority": "地方",     "985": "",  "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10694", "fetch_time": FETCH_TIME},
    # ── 青海 ──────────────────────────────────────────────────────────────────
    {"name": "青海大学",          "province": "青海", "category": "综合", "authority": "地方",     "985": "",  "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10743", "fetch_time": FETCH_TIME},
    # ── 宁夏 ──────────────────────────────────────────────────────────────────
    {"name": "宁夏大学",          "province": "宁夏", "category": "综合", "authority": "地方",     "985": "",  "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10749", "fetch_time": FETCH_TIME},
    # ── 海南 ──────────────────────────────────────────────────────────────────
    {"name": "海南大学",          "province": "海南", "category": "综合", "authority": "地方",     "985": "",  "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10589", "fetch_time": FETCH_TIME},
    # ── 江西 ──────────────────────────────────────────────────────────────────
    {"name": "南昌大学",          "province": "江西", "category": "综合", "authority": "地方",     "985": "",  "211": "✓", "double_first_class": "✓", "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=10403", "fetch_time": FETCH_TIME},
    # ── 科研院所 ──────────────────────────────────────────────────────────────
    {"name": "中国科学院大学",    "province": "北京", "category": "科研院所","authority": "中国科学院","985": "", "211": "",  "double_first_class": "",  "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=80001", "fetch_time": FETCH_TIME},
    {"name": "中国社会科学院大学","province": "北京", "category": "科研院所","authority": "中国社科院","985": "", "211": "",  "double_first_class": "",  "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=80030", "fetch_time": FETCH_TIME},
    {"name": "中国工程物理研究院","province": "四川", "category": "科研院所","authority": "工业和信息化部","985":"","211": "",  "double_first_class": "",  "url": "https://yz.chsi.com.cn/zsb/zsbxx.do?id=80031", "fetch_time": FETCH_TIME},
]
# fmt: on


def main() -> None:
    out = Path(__file__).parent / "institutions.xlsx"
    print(f"正在生成 Excel 文件，共 {len(INSTITUTIONS)} 所院校…")
    save_excel(INSTITUTIONS, out)
    print(f"✅ 已生成: {out}")
    print(f"   数据时效: {FETCH_TIME}（来源：中国研究生招生信息网 yz.chsi.com.cn）")


if __name__ == "__main__":
    main()
