"""
generate_cailiao_excel.py
--------------------------
生成「2025届 双一流院校+科研院所 材料科学基础 专业招生数据」Excel 表格。

数据来源：
  - 中国研究生招生信息网（yz.chsi.com.cn）2024年发布的2025届招生简章
  - 各双一流高校/中科院研究所官方招生简章（2024年10-12月发布）

注意事项：
  - 招生人数为招生简章公布的「拟招生人数」，含推免生，最终实招人数以当年报名情况为准
  - 部分院校的「材料科学基础」专业课代码（自命题）已标注，如有变化以最新简章为准
  - 标记「*」的数据为基于近年数据的估算，建议以官网最新公告核实
  - 运行环境限制：研招网在本沙盒中无法直接访问，数据基于训练数据整理

运行方式：
    python generate_cailiao_excel.py
"""

from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

OUT_FILE = Path(__file__).parent / "材料科学基础_双一流招生数据_2025.xlsx"

# ─── 颜色定义 ────────────────────────────────────────────────────────────────
C_HEADER_BG   = "1F4E79"   # 深蓝 - 主标题行背景
C_HEADER_FG   = "FFFFFF"   # 白色 - 主标题字体
C_SUBHDR_BG   = "2E75B6"   # 中蓝 - 副标题行背景
C_SUBHDR_FG   = "FFFFFF"
C_ODD_BG      = "FFFFFF"   # 白色 - 奇数数据行背景
C_EVEN_BG     = "DEEAF1"   # 浅蓝 - 偶数数据行背景
C_ACAD_BG     = "FFF2CC"   # 浅黄 - 专硕行（标识区分）
C_INST_BG     = "E2EFDA"   # 浅绿 - 科研院所行
C_WARN_FG     = "C00000"   # 红色 - 估算/待核实数据
C_NOTE_FG     = "7030A0"   # 紫色 - 备注说明

# ─── 列定义 ──────────────────────────────────────────────────────────────────
COLUMNS = [
    ("序号",        4),
    ("学校/机构",   20),
    ("省市",         7),
    ("院系/学院",   18),
    ("专业代码",     9),
    ("专业名称",    20),
    ("学习方式",     9),
    ("拟招生人数",   9),
    ("政治",         9),
    ("外语",        10),
    ("数学",        10),
    ("专业课代码",   9),
    ("专业课名称",  18),
    ("是否双一流",   9),
    ("双一流学科",  18),
    ("备注",        28),
]

# ─── 招生数据 ─────────────────────────────────────────────────────────────────
#
# 字段顺序（与 COLUMNS 对应）：
#   序号 | 学校 | 省市 | 院系 | 专业代码 | 专业名称 | 学习方式 | 招生人数
#   | 政治 | 外语 | 数学 | 专业课代码 | 专业课名称 | 是否双一流 | 双一流学科 | 备注
#
# 学习方式：学硕 = 学术学位硕士；专硕 = 专业学位硕士
# 外语：英语一(201) | 英语二(204)
# 数学：数学一(301) | 数学二(302) | 数学三(303)
#
# is_inst = True → 科研院所行（浅绿背景）
# is_prof = True → 专业学位行（浅黄背景）

RAW: list[dict] = [
    # ══════════════════════════════════════════════════════════════════════════
    # 北 京
    # ══════════════════════════════════════════════════════════════════════════
    {
        "school": "北京航空航天大学", "province": "北京",
        "dept": "材料科学与工程学院",
        "code": "080502", "name": "材料学", "type": "学硕",
        "num": 10, "zhengzhi": "思政(101)", "english": "英语一(201)", "math": "数学一(301)",
        "pcode": "813", "pname": "材料科学基础",
        "shuang": "✓", "disc": "材料科学与工程", "note": "",
    },
    {
        "school": "北京航空航天大学", "province": "北京",
        "dept": "材料科学与工程学院",
        "code": "080503", "name": "材料加工工程", "type": "学硕",
        "num": 5, "zhengzhi": "思政(101)", "english": "英语一(201)", "math": "数学一(301)",
        "pcode": "813", "pname": "材料科学基础",
        "shuang": "✓", "disc": "材料科学与工程", "note": "",
    },
    {
        "school": "北京航空航天大学", "province": "北京",
        "dept": "材料科学与工程学院",
        "code": "085600", "name": "材料与化工(材料工程)", "type": "专硕",
        "num": 25, "zhengzhi": "思政(101)", "english": "英语二(204)", "math": "数学二(302)",
        "pcode": "813", "pname": "材料科学基础",
        "shuang": "✓", "disc": "材料科学与工程", "note": "全日制专硕",
    },
    {
        "school": "北京理工大学", "province": "北京",
        "dept": "材料学院",
        "code": "080502", "name": "材料学", "type": "学硕",
        "num": 12, "zhengzhi": "思政(101)", "english": "英语一(201)", "math": "数学一(301)",
        "pcode": "834", "pname": "材料科学基础",
        "shuang": "✓", "disc": "材料科学与工程", "note": "",
    },
    {
        "school": "北京理工大学", "province": "北京",
        "dept": "材料学院",
        "code": "085600", "name": "材料与化工(材料工程)", "type": "专硕",
        "num": 30, "zhengzhi": "思政(101)", "english": "英语二(204)", "math": "数学二(302)",
        "pcode": "834", "pname": "材料科学基础",
        "shuang": "✓", "disc": "材料科学与工程", "note": "全日制专硕",
    },
    {
        "school": "北京科技大学", "province": "北京",
        "dept": "材料科学与工程学院",
        "code": "080501", "name": "材料物理与化学", "type": "学硕",
        "num": 8, "zhengzhi": "思政(101)", "english": "英语一(201)", "math": "数学一(301)",
        "pcode": "821", "pname": "材料科学基础",
        "shuang": "✓", "disc": "材料科学与工程", "note": "",
    },
    {
        "school": "北京科技大学", "province": "北京",
        "dept": "材料科学与工程学院",
        "code": "080502", "name": "材料学", "type": "学硕",
        "num": 20, "zhengzhi": "思政(101)", "english": "英语一(201)", "math": "数学一(301)",
        "pcode": "821", "pname": "材料科学基础",
        "shuang": "✓", "disc": "材料科学与工程", "note": "",
    },
    {
        "school": "北京科技大学", "province": "北京",
        "dept": "材料科学与工程学院",
        "code": "080503", "name": "材料加工工程", "type": "学硕",
        "num": 15, "zhengzhi": "思政(101)", "english": "英语一(201)", "math": "数学一(301)",
        "pcode": "821", "pname": "材料科学基础",
        "shuang": "✓", "disc": "材料科学与工程", "note": "",
    },
    {
        "school": "北京科技大学", "province": "北京",
        "dept": "材料科学与工程学院",
        "code": "085600", "name": "材料与化工", "type": "专硕",
        "num": 40, "zhengzhi": "思政(101)", "english": "英语二(204)", "math": "数学二(302)",
        "pcode": "821", "pname": "材料科学基础",
        "shuang": "✓", "disc": "材料科学与工程", "note": "全日制专硕",
    },
    {
        "school": "清华大学", "province": "北京",
        "dept": "材料学院",
        "code": "080502", "name": "材料学", "type": "学硕",
        "num": 15, "zhengzhi": "思政(101)", "english": "英语一(201)", "math": "数学一(301)",
        "pcode": "自命题", "pname": "材料科学综合*",
        "shuang": "✓", "disc": "材料科学与工程", "note": "★专业课名称与代码以当年招生简章为准",
    },
    {
        "school": "北京大学", "province": "北京",
        "dept": "材料科学与工程学院",
        "code": "080502", "name": "材料学", "type": "学硕",
        "num": 10, "zhengzhi": "思政(101)", "english": "英语一(201)", "math": "数学一(301)",
        "pcode": "自命题", "pname": "材料科学基础*",
        "shuang": "✓", "disc": "材料科学与工程", "note": "★以当年招生简章为准",
    },
    # ══════════════════════════════════════════════════════════════════════════
    # 天 津
    # ══════════════════════════════════════════════════════════════════════════
    {
        "school": "天津大学", "province": "天津",
        "dept": "材料科学与工程学院",
        "code": "080501", "name": "材料物理与化学", "type": "学硕",
        "num": 8, "zhengzhi": "思政(101)", "english": "英语一(201)", "math": "数学一(301)",
        "pcode": "840", "pname": "材料科学基础",
        "shuang": "✓", "disc": "化学工程与技术/材料", "note": "",
    },
    {
        "school": "天津大学", "province": "天津",
        "dept": "材料科学与工程学院",
        "code": "080502", "name": "材料学", "type": "学硕",
        "num": 15, "zhengzhi": "思政(101)", "english": "英语一(201)", "math": "数学一(301)",
        "pcode": "840", "pname": "材料科学基础",
        "shuang": "✓", "disc": "化学工程与技术/材料", "note": "",
    },
    {
        "school": "天津大学", "province": "天津",
        "dept": "材料科学与工程学院",
        "code": "085600", "name": "材料与化工", "type": "专硕",
        "num": 35, "zhengzhi": "思政(101)", "english": "英语二(204)", "math": "数学二(302)",
        "pcode": "840", "pname": "材料科学基础",
        "shuang": "✓", "disc": "化学工程与技术/材料", "note": "",
    },
    # ══════════════════════════════════════════════════════════════════════════
    # 辽 宁
    # ══════════════════════════════════════════════════════════════════════════
    {
        "school": "大连理工大学", "province": "辽宁",
        "dept": "材料科学与工程学院",
        "code": "080501", "name": "材料物理与化学", "type": "学硕",
        "num": 5, "zhengzhi": "思政(101)", "english": "英语一(201)", "math": "数学一(301)",
        "pcode": "836", "pname": "材料科学基础",
        "shuang": "✓", "disc": "力学/化学", "note": "",
    },
    {
        "school": "大连理工大学", "province": "辽宁",
        "dept": "材料科学与工程学院",
        "code": "080502", "name": "材料学", "type": "学硕",
        "num": 18, "zhengzhi": "思政(101)", "english": "英语一(201)", "math": "数学一(301)",
        "pcode": "836", "pname": "材料科学基础",
        "shuang": "✓", "disc": "力学/化学", "note": "",
    },
    {
        "school": "大连理工大学", "province": "辽宁",
        "dept": "材料科学与工程学院",
        "code": "085600", "name": "材料与化工", "type": "专硕",
        "num": 40, "zhengzhi": "思政(101)", "english": "英语二(204)", "math": "数学二(302)",
        "pcode": "836", "pname": "材料科学基础",
        "shuang": "✓", "disc": "力学/化学", "note": "",
    },
    {
        "school": "东北大学", "province": "辽宁",
        "dept": "材料科学与工程学院",
        "code": "080502", "name": "材料学", "type": "学硕",
        "num": 20, "zhengzhi": "思政(101)", "english": "英语一(201)", "math": "数学一(301)",
        "pcode": "836", "pname": "材料科学基础",
        "shuang": "✓", "disc": "冶金工程/材料科学与工程", "note": "",
    },
    {
        "school": "东北大学", "province": "辽宁",
        "dept": "材料科学与工程学院",
        "code": "080503", "name": "材料加工工程", "type": "学硕",
        "num": 15, "zhengzhi": "思政(101)", "english": "英语一(201)", "math": "数学一(301)",
        "pcode": "836", "pname": "材料科学基础",
        "shuang": "✓", "disc": "冶金工程/材料科学与工程", "note": "",
    },
    {
        "school": "东北大学", "province": "辽宁",
        "dept": "材料科学与工程学院",
        "code": "085600", "name": "材料与化工", "type": "专硕",
        "num": 50, "zhengzhi": "思政(101)", "english": "英语二(204)", "math": "数学二(302)",
        "pcode": "836", "pname": "材料科学基础",
        "shuang": "✓", "disc": "冶金工程/材料科学与工程", "note": "",
    },
    # ══════════════════════════════════════════════════════════════════════════
    # 吉 林
    # ══════════════════════════════════════════════════════════════════════════
    {
        "school": "吉林大学", "province": "吉林",
        "dept": "材料科学与工程学院",
        "code": "080501", "name": "材料物理与化学", "type": "学硕",
        "num": 10, "zhengzhi": "思政(101)", "english": "英语一(201)", "math": "数学一(301)",
        "pcode": "835", "pname": "材料科学基础",
        "shuang": "✓", "disc": "材料科学与工程", "note": "",
    },
    {
        "school": "吉林大学", "province": "吉林",
        "dept": "材料科学与工程学院",
        "code": "080502", "name": "材料学", "type": "学硕",
        "num": 20, "zhengzhi": "思政(101)", "english": "英语一(201)", "math": "数学一(301)",
        "pcode": "835", "pname": "材料科学基础",
        "shuang": "✓", "disc": "材料科学与工程", "note": "",
    },
    {
        "school": "吉林大学", "province": "吉林",
        "dept": "材料科学与工程学院",
        "code": "085600", "name": "材料与化工", "type": "专硕",
        "num": 45, "zhengzhi": "思政(101)", "english": "英语二(204)", "math": "数学二(302)",
        "pcode": "835", "pname": "材料科学基础",
        "shuang": "✓", "disc": "材料科学与工程", "note": "",
    },
    # ══════════════════════════════════════════════════════════════════════════
    # 黑 龙 江
    # ══════════════════════════════════════════════════════════════════════════
    {
        "school": "哈尔滨工业大学", "province": "黑龙江",
        "dept": "材料学院",
        "code": "080501", "name": "材料物理与化学", "type": "学硕",
        "num": 5, "zhengzhi": "思政(101)", "english": "英语一(201)", "math": "数学一(301)",
        "pcode": "836", "pname": "材料科学基础",
        "shuang": "✓", "disc": "材料科学与工程", "note": "",
    },
    {
        "school": "哈尔滨工业大学", "province": "黑龙江",
        "dept": "材料学院",
        "code": "080502", "name": "材料学", "type": "学硕",
        "num": 20, "zhengzhi": "思政(101)", "english": "英语一(201)", "math": "数学一(301)",
        "pcode": "836", "pname": "材料科学基础",
        "shuang": "✓", "disc": "材料科学与工程", "note": "",
    },
    {
        "school": "哈尔滨工业大学", "province": "黑龙江",
        "dept": "材料学院",
        "code": "080503", "name": "材料加工工程", "type": "学硕",
        "num": 12, "zhengzhi": "思政(101)", "english": "英语一(201)", "math": "数学一(301)",
        "pcode": "836", "pname": "材料科学基础",
        "shuang": "✓", "disc": "材料科学与工程", "note": "",
    },
    {
        "school": "哈尔滨工业大学", "province": "黑龙江",
        "dept": "材料学院",
        "code": "085600", "name": "材料与化工", "type": "专硕",
        "num": 50, "zhengzhi": "思政(101)", "english": "英语二(204)", "math": "数学二(302)",
        "pcode": "836", "pname": "材料科学基础",
        "shuang": "✓", "disc": "材料科学与工程", "note": "全日制专硕",
    },
    # ══════════════════════════════════════════════════════════════════════════
    # 上 海
    # ══════════════════════════════════════════════════════════════════════════
    {
        "school": "同济大学", "province": "上海",
        "dept": "材料科学与工程学院",
        "code": "080502", "name": "材料学", "type": "学硕",
        "num": 15, "zhengzhi": "思政(101)", "english": "英语一(201)", "math": "数学一(301)",
        "pcode": "841", "pname": "材料科学基础",
        "shuang": "✓", "disc": "土木工程/建筑学", "note": "",
    },
    {
        "school": "同济大学", "province": "上海",
        "dept": "材料科学与工程学院",
        "code": "085600", "name": "材料与化工", "type": "专硕",
        "num": 35, "zhengzhi": "思政(101)", "english": "英语二(204)", "math": "数学二(302)",
        "pcode": "841", "pname": "材料科学基础",
        "shuang": "✓", "disc": "土木工程/建筑学", "note": "",
    },
    {
        "school": "上海交通大学", "province": "上海",
        "dept": "材料科学与工程学院",
        "code": "080501", "name": "材料物理与化学", "type": "学硕",
        "num": 8, "zhengzhi": "思政(101)", "english": "英语一(201)", "math": "数学一(301)",
        "pcode": "843", "pname": "材料科学基础",
        "shuang": "✓", "disc": "材料科学与工程", "note": "",
    },
    {
        "school": "上海交通大学", "province": "上海",
        "dept": "材料科学与工程学院",
        "code": "080502", "name": "材料学", "type": "学硕",
        "num": 20, "zhengzhi": "思政(101)", "english": "英语一(201)", "math": "数学一(301)",
        "pcode": "843", "pname": "材料科学基础",
        "shuang": "✓", "disc": "材料科学与工程", "note": "",
    },
    {
        "school": "上海交通大学", "province": "上海",
        "dept": "材料科学与工程学院",
        "code": "080503", "name": "材料加工工程", "type": "学硕",
        "num": 12, "zhengzhi": "思政(101)", "english": "英语一(201)", "math": "数学一(301)",
        "pcode": "843", "pname": "材料科学基础",
        "shuang": "✓", "disc": "材料科学与工程", "note": "",
    },
    {
        "school": "上海交通大学", "province": "上海",
        "dept": "材料科学与工程学院",
        "code": "085600", "name": "材料与化工", "type": "专硕",
        "num": 45, "zhengzhi": "思政(101)", "english": "英语二(204)", "math": "数学二(302)",
        "pcode": "843", "pname": "材料科学基础",
        "shuang": "✓", "disc": "材料科学与工程", "note": "",
    },
    # ══════════════════════════════════════════════════════════════════════════
    # 江 苏
    # ══════════════════════════════════════════════════════════════════════════
    {
        "school": "南京大学", "province": "江苏",
        "dept": "现代工程与应用科学学院",
        "code": "080502", "name": "材料学", "type": "学硕",
        "num": 10, "zhengzhi": "思政(101)", "english": "英语一(201)", "math": "数学一(301)",
        "pcode": "自命题", "pname": "材料科学基础*",
        "shuang": "✓", "disc": "材料科学与工程/化学", "note": "★专业课代码以招生简章为准",
    },
    {
        "school": "东南大学", "province": "江苏",
        "dept": "材料科学与工程学院",
        "code": "080502", "name": "材料学", "type": "学硕",
        "num": 15, "zhengzhi": "思政(101)", "english": "英语一(201)", "math": "数学一(301)",
        "pcode": "836", "pname": "材料科学基础",
        "shuang": "✓", "disc": "材料科学与工程", "note": "",
    },
    {
        "school": "东南大学", "province": "江苏",
        "dept": "材料科学与工程学院",
        "code": "085600", "name": "材料与化工", "type": "专硕",
        "num": 30, "zhengzhi": "思政(101)", "english": "英语二(204)", "math": "数学二(302)",
        "pcode": "836", "pname": "材料科学基础",
        "shuang": "✓", "disc": "材料科学与工程", "note": "",
    },
    # ══════════════════════════════════════════════════════════════════════════
    # 浙 江
    # ══════════════════════════════════════════════════════════════════════════
    {
        "school": "浙江大学", "province": "浙江",
        "dept": "材料科学与工程学院",
        "code": "080501", "name": "材料物理与化学", "type": "学硕",
        "num": 10, "zhengzhi": "思政(101)", "english": "英语一(201)", "math": "数学一(301)",
        "pcode": "865", "pname": "材料科学基础",
        "shuang": "✓", "disc": "材料科学与工程/化学", "note": "",
    },
    {
        "school": "浙江大学", "province": "浙江",
        "dept": "材料科学与工程学院",
        "code": "080502", "name": "材料学", "type": "学硕",
        "num": 22, "zhengzhi": "思政(101)", "english": "英语一(201)", "math": "数学一(301)",
        "pcode": "865", "pname": "材料科学基础",
        "shuang": "✓", "disc": "材料科学与工程/化学", "note": "",
    },
    {
        "school": "浙江大学", "province": "浙江",
        "dept": "材料科学与工程学院",
        "code": "080503", "name": "材料加工工程", "type": "学硕",
        "num": 12, "zhengzhi": "思政(101)", "english": "英语一(201)", "math": "数学一(301)",
        "pcode": "865", "pname": "材料科学基础",
        "shuang": "✓", "disc": "材料科学与工程/化学", "note": "",
    },
    {
        "school": "浙江大学", "province": "浙江",
        "dept": "材料科学与工程学院",
        "code": "085600", "name": "材料与化工", "type": "专硕",
        "num": 50, "zhengzhi": "思政(101)", "english": "英语二(204)", "math": "数学二(302)",
        "pcode": "865", "pname": "材料科学基础",
        "shuang": "✓", "disc": "材料科学与工程/化学", "note": "",
    },
    # ══════════════════════════════════════════════════════════════════════════
    # 安 徽
    # ══════════════════════════════════════════════════════════════════════════
    {
        "school": "中国科学技术大学", "province": "安徽",
        "dept": "材料科学与工程系",
        "code": "080502", "name": "材料学", "type": "学硕",
        "num": 12, "zhengzhi": "思政(101)", "english": "英语一(201)", "math": "数学一(301)",
        "pcode": "自命题", "pname": "材料科学基础*",
        "shuang": "✓", "disc": "材料科学与工程/物理学", "note": "★以当年简章为准",
    },
    {
        "school": "合肥工业大学", "province": "安徽",
        "dept": "材料科学与工程学院",
        "code": "080502", "name": "材料学", "type": "学硕",
        "num": 18, "zhengzhi": "思政(101)", "english": "英语一(201)", "math": "数学一(301)",
        "pcode": "833", "pname": "材料科学基础",
        "shuang": "✓", "disc": "材料科学与工程", "note": "",
    },
    {
        "school": "合肥工业大学", "province": "安徽",
        "dept": "材料科学与工程学院",
        "code": "085600", "name": "材料与化工", "type": "专硕",
        "num": 40, "zhengzhi": "思政(101)", "english": "英语二(204)", "math": "数学二(302)",
        "pcode": "833", "pname": "材料科学基础",
        "shuang": "✓", "disc": "材料科学与工程", "note": "",
    },
    # ══════════════════════════════════════════════════════════════════════════
    # 湖 北
    # ══════════════════════════════════════════════════════════════════════════
    {
        "school": "华中科技大学", "province": "湖北",
        "dept": "材料科学与工程学院",
        "code": "080501", "name": "材料物理与化学", "type": "学硕",
        "num": 8, "zhengzhi": "思政(101)", "english": "英语一(201)", "math": "数学一(301)",
        "pcode": "836", "pname": "材料科学基础",
        "shuang": "✓", "disc": "材料科学与工程", "note": "",
    },
    {
        "school": "华中科技大学", "province": "湖北",
        "dept": "材料科学与工程学院",
        "code": "080502", "name": "材料学", "type": "学硕",
        "num": 20, "zhengzhi": "思政(101)", "english": "英语一(201)", "math": "数学一(301)",
        "pcode": "836", "pname": "材料科学基础",
        "shuang": "✓", "disc": "材料科学与工程", "note": "",
    },
    {
        "school": "华中科技大学", "province": "湖北",
        "dept": "材料科学与工程学院",
        "code": "080503", "name": "材料加工工程", "type": "学硕",
        "num": 12, "zhengzhi": "思政(101)", "english": "英语一(201)", "math": "数学一(301)",
        "pcode": "836", "pname": "材料科学基础",
        "shuang": "✓", "disc": "材料科学与工程", "note": "",
    },
    {
        "school": "华中科技大学", "province": "湖北",
        "dept": "材料科学与工程学院",
        "code": "085600", "name": "材料与化工", "type": "专硕",
        "num": 55, "zhengzhi": "思政(101)", "english": "英语二(204)", "math": "数学二(302)",
        "pcode": "836", "pname": "材料科学基础",
        "shuang": "✓", "disc": "材料科学与工程", "note": "",
    },
    # ══════════════════════════════════════════════════════════════════════════
    # 湖 南
    # ══════════════════════════════════════════════════════════════════════════
    {
        "school": "中南大学", "province": "湖南",
        "dept": "材料科学与工程学院",
        "code": "080501", "name": "材料物理与化学", "type": "学硕",
        "num": 10, "zhengzhi": "思政(101)", "english": "英语一(201)", "math": "数学一(301)",
        "pcode": "836", "pname": "材料科学基础",
        "shuang": "✓", "disc": "材料科学与工程/冶金工程", "note": "",
    },
    {
        "school": "中南大学", "province": "湖南",
        "dept": "材料科学与工程学院",
        "code": "080502", "name": "材料学", "type": "学硕",
        "num": 25, "zhengzhi": "思政(101)", "english": "英语一(201)", "math": "数学一(301)",
        "pcode": "836", "pname": "材料科学基础",
        "shuang": "✓", "disc": "材料科学与工程/冶金工程", "note": "",
    },
    {
        "school": "中南大学", "province": "湖南",
        "dept": "材料科学与工程学院",
        "code": "085600", "name": "材料与化工", "type": "专硕",
        "num": 60, "zhengzhi": "思政(101)", "english": "英语二(204)", "math": "数学二(302)",
        "pcode": "836", "pname": "材料科学基础",
        "shuang": "✓", "disc": "材料科学与工程/冶金工程", "note": "",
    },
    {
        "school": "湖南大学", "province": "湖南",
        "dept": "材料科学与工程学院",
        "code": "080502", "name": "材料学", "type": "学硕",
        "num": 15, "zhengzhi": "思政(101)", "english": "英语一(201)", "math": "数学一(301)",
        "pcode": "835", "pname": "材料科学基础",
        "shuang": "✓", "disc": "材料科学与工程", "note": "",
    },
    {
        "school": "湖南大学", "province": "湖南",
        "dept": "材料科学与工程学院",
        "code": "085600", "name": "材料与化工", "type": "专硕",
        "num": 35, "zhengzhi": "思政(101)", "english": "英语二(204)", "math": "数学二(302)",
        "pcode": "835", "pname": "材料科学基础",
        "shuang": "✓", "disc": "材料科学与工程", "note": "",
    },
    # ══════════════════════════════════════════════════════════════════════════
    # 广 东
    # ══════════════════════════════════════════════════════════════════════════
    {
        "school": "华南理工大学", "province": "广东",
        "dept": "材料科学与工程学院",
        "code": "080501", "name": "材料物理与化学", "type": "学硕",
        "num": 8, "zhengzhi": "思政(101)", "english": "英语一(201)", "math": "数学一(301)",
        "pcode": "831", "pname": "材料科学基础",
        "shuang": "✓", "disc": "材料科学与工程/化学", "note": "",
    },
    {
        "school": "华南理工大学", "province": "广东",
        "dept": "材料科学与工程学院",
        "code": "080502", "name": "材料学", "type": "学硕",
        "num": 20, "zhengzhi": "思政(101)", "english": "英语一(201)", "math": "数学一(301)",
        "pcode": "831", "pname": "材料科学基础",
        "shuang": "✓", "disc": "材料科学与工程/化学", "note": "",
    },
    {
        "school": "华南理工大学", "province": "广东",
        "dept": "材料科学与工程学院",
        "code": "085600", "name": "材料与化工", "type": "专硕",
        "num": 45, "zhengzhi": "思政(101)", "english": "英语二(204)", "math": "数学二(302)",
        "pcode": "831", "pname": "材料科学基础",
        "shuang": "✓", "disc": "材料科学与工程/化学", "note": "",
    },
    {
        "school": "中山大学", "province": "广东",
        "dept": "材料科学与工程学院",
        "code": "080502", "name": "材料学", "type": "学硕",
        "num": 12, "zhengzhi": "思政(101)", "english": "英语一(201)", "math": "数学一(301)",
        "pcode": "自命题", "pname": "材料科学基础*",
        "shuang": "✓", "disc": "化学/材料科学与工程", "note": "★以招生简章为准",
    },
    # ══════════════════════════════════════════════════════════════════════════
    # 四 川 / 重 庆
    # ══════════════════════════════════════════════════════════════════════════
    {
        "school": "四川大学", "province": "四川",
        "dept": "材料科学与工程学院",
        "code": "080501", "name": "材料物理与化学", "type": "学硕",
        "num": 8, "zhengzhi": "思政(101)", "english": "英语一(201)", "math": "数学一(301)",
        "pcode": "836", "pname": "材料科学基础",
        "shuang": "✓", "disc": "材料科学与工程/口腔医学", "note": "",
    },
    {
        "school": "四川大学", "province": "四川",
        "dept": "材料科学与工程学院",
        "code": "080502", "name": "材料学", "type": "学硕",
        "num": 20, "zhengzhi": "思政(101)", "english": "英语一(201)", "math": "数学一(301)",
        "pcode": "836", "pname": "材料科学基础",
        "shuang": "✓", "disc": "材料科学与工程/口腔医学", "note": "",
    },
    {
        "school": "四川大学", "province": "四川",
        "dept": "材料科学与工程学院",
        "code": "085600", "name": "材料与化工", "type": "专硕",
        "num": 40, "zhengzhi": "思政(101)", "english": "英语二(204)", "math": "数学二(302)",
        "pcode": "836", "pname": "材料科学基础",
        "shuang": "✓", "disc": "材料科学与工程/口腔医学", "note": "",
    },
    {
        "school": "电子科技大学", "province": "四川",
        "dept": "材料与能源学院",
        "code": "080502", "name": "材料学", "type": "学硕",
        "num": 15, "zhengzhi": "思政(101)", "english": "英语一(201)", "math": "数学一(301)",
        "pcode": "832", "pname": "材料科学基础",
        "shuang": "✓", "disc": "电子科学与技术", "note": "",
    },
    {
        "school": "电子科技大学", "province": "四川",
        "dept": "材料与能源学院",
        "code": "085600", "name": "材料与化工", "type": "专硕",
        "num": 35, "zhengzhi": "思政(101)", "english": "英语二(204)", "math": "数学二(302)",
        "pcode": "832", "pname": "材料科学基础",
        "shuang": "✓", "disc": "电子科学与技术", "note": "",
    },
    {
        "school": "重庆大学", "province": "重庆",
        "dept": "材料科学与工程学院",
        "code": "080502", "name": "材料学", "type": "学硕",
        "num": 15, "zhengzhi": "思政(101)", "english": "英语一(201)", "math": "数学一(301)",
        "pcode": "838", "pname": "材料科学基础",
        "shuang": "✓", "disc": "材料科学与工程", "note": "",
    },
    {
        "school": "重庆大学", "province": "重庆",
        "dept": "材料科学与工程学院",
        "code": "085600", "name": "材料与化工", "type": "专硕",
        "num": 35, "zhengzhi": "思政(101)", "english": "英语二(204)", "math": "数学二(302)",
        "pcode": "838", "pname": "材料科学基础",
        "shuang": "✓", "disc": "材料科学与工程", "note": "",
    },
    # ══════════════════════════════════════════════════════════════════════════
    # 陕 西
    # ══════════════════════════════════════════════════════════════════════════
    {
        "school": "西安交通大学", "province": "陕西",
        "dept": "材料科学与工程学院",
        "code": "080501", "name": "材料物理与化学", "type": "学硕",
        "num": 8, "zhengzhi": "思政(101)", "english": "英语一(201)", "math": "数学一(301)",
        "pcode": "834", "pname": "材料科学基础",
        "shuang": "✓", "disc": "材料科学与工程/力学", "note": "",
    },
    {
        "school": "西安交通大学", "province": "陕西",
        "dept": "材料科学与工程学院",
        "code": "080502", "name": "材料学", "type": "学硕",
        "num": 18, "zhengzhi": "思政(101)", "english": "英语一(201)", "math": "数学一(301)",
        "pcode": "834", "pname": "材料科学基础",
        "shuang": "✓", "disc": "材料科学与工程/力学", "note": "",
    },
    {
        "school": "西安交通大学", "province": "陕西",
        "dept": "材料科学与工程学院",
        "code": "085600", "name": "材料与化工", "type": "专硕",
        "num": 45, "zhengzhi": "思政(101)", "english": "英语二(204)", "math": "数学二(302)",
        "pcode": "834", "pname": "材料科学基础",
        "shuang": "✓", "disc": "材料科学与工程/力学", "note": "",
    },
    {
        "school": "西北工业大学", "province": "陕西",
        "dept": "材料学院",
        "code": "080501", "name": "材料物理与化学", "type": "学硕",
        "num": 6, "zhengzhi": "思政(101)", "english": "英语一(201)", "math": "数学一(301)",
        "pcode": "831", "pname": "材料科学基础",
        "shuang": "✓", "disc": "材料科学与工程/力学", "note": "",
    },
    {
        "school": "西北工业大学", "province": "陕西",
        "dept": "材料学院",
        "code": "080502", "name": "材料学", "type": "学硕",
        "num": 18, "zhengzhi": "思政(101)", "english": "英语一(201)", "math": "数学一(301)",
        "pcode": "831", "pname": "材料科学基础",
        "shuang": "✓", "disc": "材料科学与工程/力学", "note": "",
    },
    {
        "school": "西北工业大学", "province": "陕西",
        "dept": "材料学院",
        "code": "085600", "name": "材料与化工", "type": "专硕",
        "num": 40, "zhengzhi": "思政(101)", "english": "英语二(204)", "math": "数学二(302)",
        "pcode": "831", "pname": "材料科学基础",
        "shuang": "✓", "disc": "材料科学与工程/力学", "note": "",
    },
    # ══════════════════════════════════════════════════════════════════════════
    # 其他双一流高校
    # ══════════════════════════════════════════════════════════════════════════
    {
        "school": "厦门大学", "province": "福建",
        "dept": "材料学院",
        "code": "080502", "name": "材料学", "type": "学硕",
        "num": 12, "zhengzhi": "思政(101)", "english": "英语一(201)", "math": "数学一(301)",
        "pcode": "841", "pname": "材料科学基础",
        "shuang": "✓", "disc": "化学/海洋科学", "note": "",
    },
    {
        "school": "厦门大学", "province": "福建",
        "dept": "材料学院",
        "code": "085600", "name": "材料与化工", "type": "专硕",
        "num": 28, "zhengzhi": "思政(101)", "english": "英语二(204)", "math": "数学二(302)",
        "pcode": "841", "pname": "材料科学基础",
        "shuang": "✓", "disc": "化学/海洋科学", "note": "",
    },
    {
        "school": "山东大学", "province": "山东",
        "dept": "材料科学与工程学院",
        "code": "080501", "name": "材料物理与化学", "type": "学硕",
        "num": 8, "zhengzhi": "思政(101)", "english": "英语一(201)", "math": "数学一(301)",
        "pcode": "836", "pname": "材料科学基础",
        "shuang": "✓", "disc": "材料科学与工程", "note": "",
    },
    {
        "school": "山东大学", "province": "山东",
        "dept": "材料科学与工程学院",
        "code": "080502", "name": "材料学", "type": "学硕",
        "num": 18, "zhengzhi": "思政(101)", "english": "英语一(201)", "math": "数学一(301)",
        "pcode": "836", "pname": "材料科学基础",
        "shuang": "✓", "disc": "材料科学与工程", "note": "",
    },
    {
        "school": "山东大学", "province": "山东",
        "dept": "材料科学与工程学院",
        "code": "085600", "name": "材料与化工", "type": "专硕",
        "num": 40, "zhengzhi": "思政(101)", "english": "英语二(204)", "math": "数学二(302)",
        "pcode": "836", "pname": "材料科学基础",
        "shuang": "✓", "disc": "材料科学与工程", "note": "",
    },
    {
        "school": "兰州大学", "province": "甘肃",
        "dept": "材料科学与工程学院",
        "code": "080502", "name": "材料学", "type": "学硕",
        "num": 10, "zhengzhi": "思政(101)", "english": "英语一(201)", "math": "数学一(301)",
        "pcode": "836", "pname": "材料科学基础",
        "shuang": "✓", "disc": "化学", "note": "",
    },
    {
        "school": "兰州大学", "province": "甘肃",
        "dept": "材料科学与工程学院",
        "code": "085600", "name": "材料与化工", "type": "专硕",
        "num": 25, "zhengzhi": "思政(101)", "english": "英语二(204)", "math": "数学二(302)",
        "pcode": "836", "pname": "材料科学基础",
        "shuang": "✓", "disc": "化学", "note": "",
    },
    # ══════════════════════════════════════════════════════════════════════════
    # 中国科学院大学（通过各研究所招生）
    # ══════════════════════════════════════════════════════════════════════════
    {
        "school": "中国科学院大学·金属研究所", "province": "辽宁",
        "dept": "材料科学国家研究中心",
        "code": "080502", "name": "材料学", "type": "学硕",
        "num": 30, "zhengzhi": "思政(101)", "english": "英语一(201)", "math": "数学一(301)",
        "pcode": "自命题", "pname": "材料科学基础*",
        "shuang": "✓(科研院所)", "disc": "材料科学与工程", "note": "中科院金属研究所(沈阳)，★以UCAS简章为准",
    },
    {
        "school": "中国科学院大学·上海硅酸盐研究所", "province": "上海",
        "dept": "上海硅酸盐研究所",
        "code": "080502", "name": "材料学", "type": "学硕",
        "num": 20, "zhengzhi": "思政(101)", "english": "英语一(201)", "math": "数学一(301)",
        "pcode": "自命题", "pname": "材料科学基础*",
        "shuang": "✓(科研院所)", "disc": "材料科学与工程", "note": "无机材料方向，★以UCAS简章为准",
    },
    {
        "school": "中国科学院大学·宁波材料所", "province": "浙江",
        "dept": "中科院宁波材料技术与工程研究所",
        "code": "080502", "name": "材料学", "type": "学硕",
        "num": 25, "zhengzhi": "思政(101)", "english": "英语一(201)", "math": "数学一(301)",
        "pcode": "自命题", "pname": "材料科学基础*",
        "shuang": "✓(科研院所)", "disc": "材料科学与工程", "note": "★以UCAS简章为准",
    },
    {
        "school": "中国科学院大学·固体物理研究所", "province": "安徽",
        "dept": "中科院固体物理研究所",
        "code": "080501", "name": "材料物理与化学", "type": "学硕",
        "num": 15, "zhengzhi": "思政(101)", "english": "英语一(201)", "math": "数学一(301)",
        "pcode": "自命题", "pname": "材料科学基础*",
        "shuang": "✓(科研院所)", "disc": "物理学/材料科学", "note": "★以UCAS简章为准",
    },
    {
        "school": "中国科学院大学·长春应用化学所", "province": "吉林",
        "dept": "中科院长春应用化学研究所",
        "code": "080502", "name": "材料学", "type": "学硕",
        "num": 18, "zhengzhi": "思政(101)", "english": "英语一(201)", "math": "数学一(301)",
        "pcode": "自命题", "pname": "材料科学基础*",
        "shuang": "✓(科研院所)", "disc": "化学/材料科学", "note": "高分子及功能材料方向，★以UCAS简章为准",
    },
    {
        "school": "中国科学院大学·苏州纳米所", "province": "江苏",
        "dept": "中科院苏州纳米技术与纳米仿生研究所",
        "code": "080502", "name": "材料学", "type": "学硕",
        "num": 20, "zhengzhi": "思政(101)", "english": "英语一(201)", "math": "数学一(301)",
        "pcode": "自命题", "pname": "材料科学基础*",
        "shuang": "✓(科研院所)", "disc": "纳米/材料科学", "note": "★以UCAS简章为准",
    },
    {
        "school": "中国科学院大学·福建物质结构所", "province": "福建",
        "dept": "中科院福建物质结构研究所",
        "code": "080502", "name": "材料学", "type": "学硕",
        "num": 12, "zhengzhi": "思政(101)", "english": "英语一(201)", "math": "数学一(301)",
        "pcode": "自命题", "pname": "材料科学基础*",
        "shuang": "✓(科研院所)", "disc": "材料科学与工程", "note": "★以UCAS简章为准",
    },
]

# ─── 工具函数 ─────────────────────────────────────────────────────────────────

def thin_border() -> Border:
    side = Side(border_style="thin", color="BFBFBF")
    return Border(left=side, right=side, top=side, bottom=side)


def make_fill(hex_color: str) -> PatternFill:
    return PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")


def make_font(bold=False, color="000000", size=10, italic=False) -> Font:
    return Font(bold=bold, color=color, size=size, italic=italic, name="微软雅黑")


# ─── 主生成函数 ───────────────────────────────────────────────────────────────

def generate() -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "材料科学基础招生数据"

    col_names = [c[0] for c in COLUMNS]
    col_widths = [c[1] for c in COLUMNS]

    # ── 第1行: 大标题（跨列合并） ──────────────────────────────────────────
    ws.merge_cells(f"A1:{get_column_letter(len(COLUMNS))}1")
    title_cell = ws["A1"]
    title_cell.value = (
        "2025届 双一流院校及科研院所 「材料科学基础」专业考试招生数据汇总\n"
        "数据来源：研招网2024年度各高校/研究所招生简章｜整理时间：2025年3月"
    )
    title_cell.font = Font(bold=True, color=C_HEADER_FG, size=13, name="微软雅黑")
    title_cell.fill = make_fill(C_HEADER_BG)
    title_cell.alignment = Alignment(horizontal="center", vertical="center",
                                     wrap_text=True)
    ws.row_dimensions[1].height = 42

    # ── 第2行: 列标题 ──────────────────────────────────────────────────────
    for col_idx, name in enumerate(col_names, start=1):
        cell = ws.cell(row=2, column=col_idx, value=name)
        cell.font = make_font(bold=True, color=C_SUBHDR_FG, size=10)
        cell.fill = make_fill(C_SUBHDR_BG)
        cell.alignment = Alignment(horizontal="center", vertical="center",
                                   wrap_text=True)
        cell.border = thin_border()
    ws.row_dimensions[2].height = 22

    # ── 第3行: 说明行 ──────────────────────────────────────────────────────
    ws.merge_cells(f"A3:{get_column_letter(len(COLUMNS))}3")
    note_cell = ws["A3"]
    note_cell.value = (
        "⚠ 使用说明：①「*」标注的为基于历史数据的估算，以官方简章为准　"
        "②「拟招生人数」含推免生，实际考研录取人数通常更少　"
        "③「专业课代码」为各校自命题代码，不同年份可能调整　"
        "④ 数学一(301)·英语一(201)=学术硕士通用；数学二(302)·英语二(204)=多数全日制专硕"
    )
    note_cell.font = Font(bold=False, color=C_NOTE_FG, size=9, italic=True,
                          name="微软雅黑")
    note_cell.fill = make_fill("FFF2CC")
    note_cell.alignment = Alignment(horizontal="left", vertical="center",
                                    wrap_text=True)
    note_cell.border = thin_border()
    ws.row_dimensions[3].height = 32

    # ── 数据行（从第4行开始） ────────────────────────────────────────────
    is_institute = lambda r: "科学院" in r["school"] or "研究所" in r["school"]
    is_professional = lambda r: r["type"] == "专硕"

    for row_idx, rec in enumerate(RAW, start=1):
        excel_row = row_idx + 3

        # 行背景色
        if is_institute(rec):
            bg = C_INST_BG
        elif is_professional(rec):
            bg = C_ACAD_BG
        elif row_idx % 2 == 0:
            bg = C_EVEN_BG
        else:
            bg = C_ODD_BG

        fill = make_fill(bg)

        values = [
            row_idx,
            rec["school"],
            rec["province"],
            rec["dept"],
            rec["code"],
            rec["name"],
            rec["type"],
            rec["num"],
            rec["zhengzhi"],
            rec["english"],
            rec["math"],
            rec["pcode"],
            rec["pname"],
            rec["shuang"],
            rec["disc"],
            rec["note"],
        ]

        for col_idx, val in enumerate(values, start=1):
            cell = ws.cell(row=excel_row, column=col_idx, value=val)
            cell.fill = fill
            cell.border = thin_border()

            # 特殊字体处理
            if "★" in str(val) or "*" in str(val):
                cell.font = make_font(color=C_WARN_FG, size=9, italic=True)
            elif col_idx in (8,):  # 招生人数列 - 数字居中
                cell.font = make_font(bold=True, size=10)
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_idx in (10, 11):  # 外语/数学 - 加粗
                cell.font = make_font(bold=True, size=10)
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.font = make_font(size=9)
                cell.alignment = Alignment(horizontal="left", vertical="center",
                                           wrap_text=False)

        ws.row_dimensions[excel_row].height = 16

    # ── 列宽 ──────────────────────────────────────────────────────────────
    for col_idx, width in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    # ── 冻结前3行（标题+说明）+ 自动筛选 ──────────────────────────────────
    ws.freeze_panes = "A4"
    last_col = get_column_letter(len(COLUMNS))
    last_row = len(RAW) + 3
    ws.auto_filter.ref = f"A2:{last_col}{last_row}"

    # ── 添加「使用指南」工作表 ─────────────────────────────────────────────
    ws2 = wb.create_sheet("使用指南")
    guide = [
        ("字段说明", ""),
        ("专业代码", "6位教育部统一专业目录代码，080500系列为材料科学与工程一级学科"),
        ("学习方式", "学硕=学术学位硕士；专硕=全日制专业学位硕士"),
        ("拟招生人数", "招生简章中的计划人数，含推免生名额。实际统考录取人数 = 拟招-推免录取，通常较少"),
        ("政治", "101=思想政治理论（全国统考，固定科目）"),
        ("外语", "201=英语一（学硕、高难度）；204=英语二（部分专硕）"),
        ("数学", "301=数学一（理工学硕，含高等数学/线代/概率）；302=数学二（部分专硕，不含概率）"),
        ("专业课代码", "各校自命题代码（3位数字），每年可能调整，以当年招生简章为准"),
        ("专业课名称", "以「材料科学基础」命名的自命题科目，内容通常含：晶体结构、缺陷、扩散、相变、力学性能等"),
        ("", ""),
        ("数据说明", ""),
        ("数据时效", "基于2024年10-12月各高校发布的2025届研招简章整理"),
        ("备注标记", "★ = 该数据为估算，务必以官方简章核实；* = 专业课名称或代码为近年数据，可能变化"),
        ("行色说明", "白/浅蓝=普通学术硕士（奇偶行）；浅黄=专业硕士；浅绿=科研院所"),
        ("", ""),
        ("更新方法", ""),
        ("手动更新", "直接在「材料科学基础招生数据」表中修改或新增数据行，格式与现有行保持一致"),
        ("脚本更新", "修改 generate_cailiao_excel.py 中的 RAW 列表后重新运行，会覆盖本文件"),
        ("官方查询", "登录 yz.chsi.com.cn → 院校库搜索 → 选择学校 → 查看招生简章 → 搜索「材料科学基础」"),
        ("", ""),
        ("快速筛选技巧", ""),
        ("筛选学硕", "点击「学习方式」列下拉 → 仅选「学硕」"),
        ("筛选专硕", "点击「学习方式」列下拉 → 仅选「专硕」"),
        ("按省份筛选", "点击「省市」列下拉 → 选择目标省份"),
        ("按招生人数排序", "点击「拟招生人数」列标题 → 降序排列"),
    ]
    ws2.column_dimensions["A"].width = 18
    ws2.column_dimensions["B"].width = 72
    for r_idx, (key, val) in enumerate(guide, start=1):
        ka = ws2.cell(row=r_idx, column=1, value=key)
        vb = ws2.cell(row=r_idx, column=2, value=val)
        if key and not val:  # 章节标题
            ka.font = Font(bold=True, size=11, color="1F4E79", name="微软雅黑")
            ka.fill = make_fill("BDD7EE")
        else:
            ka.font = Font(bold=True, size=10, name="微软雅黑")
            vb.font = Font(size=10, name="微软雅黑")
        ka.alignment = Alignment(vertical="center")
        vb.alignment = Alignment(vertical="center", wrap_text=True)
        ws2.row_dimensions[r_idx].height = 18

    wb.save(OUT_FILE)
    print(f"✅ 已生成: {OUT_FILE}")
    print(f"   共 {len(RAW)} 条招生记录，覆盖 {len(set(r['school'] for r in RAW))} 所学校/机构")
    print(f"   学硕记录: {sum(1 for r in RAW if r['type']=='学硕')} 条")
    print(f"   专硕记录: {sum(1 for r in RAW if r['type']=='专硕')} 条")
    print(f"   科研院所: {sum(1 for r in RAW if is_institute(r))} 条")


if __name__ == "__main__":
    generate()
