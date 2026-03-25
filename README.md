# zsh_code_test

## 研招网院校数据爬取工具

使用 Python 从[中国研究生招生信息网（研招网）](https://yz.chsi.com.cn)拉取招生院校数据，结果保存为格式化 Excel（.xlsx）、CSV 和 JSON 文件。

### 📊 已生成数据文件

**[`institutions.xlsx`](./institutions.xlsx)** — 开箱即用的 Excel 表格，已收录 **144 所**具有研究生招生资格的高等院校及科研机构（数据时效：2025-2026 学年）。

表格包含以下字段：

| 列 | 说明 |
|----|------|
| 序号 | 自动编号 |
| 院校名称 | 招生单位全称 |
| 所在省份 | 院校所在省/市/自治区 |
| 院校类型 | 综合/理工/师范/医药/农林/财经/政法/民族/艺术/体育/语言/科研院所 |
| 管理部门 | 教育部直属/其他部委/地方 |
| 985工程 | 是否 985 工程院校（✓ 表示是） |
| 211工程 | 是否 211 工程院校（✓ 表示是） |
| 双一流 | 是否双一流建设高校（✓ 表示是） |
| 研招网链接 | 院校在研招网的招生简章链接（可点击） |
| 数据更新时间 | 数据对应学年 |

表格支持**排序、筛选、编辑**，首两行（中英文标题）已冻结，可向下滚动浏览全部数据。

### 功能

- 从研招网实时爬取全国或指定省份的招生院校列表
- 解析院校名称、省份、类别及详情链接
- 自动翻页，支持最大页数限制
- 保存为格式化 **Excel**（.xlsx）、CSV 和 JSON 文件

### 环境要求

- Python 3.10+

### 安装依赖

```bash
pip install -r requirements.txt
```

### 使用方法

#### 方式一：实时爬取（需能访问 yz.chsi.com.cn）

```bash
# 爬取全国院校（默认最多 50 页，每页间隔 1 秒）
python scraper.py

# 按省份过滤
python scraper.py --province 北京

# 自定义参数
python scraper.py --province 上海 --max-pages 10 --delay 1.5 --output-dir ./data
```

**参数说明**

| 参数 | 默认值 | 说明 |
|------|--------|------|
| `--province` | 全国 | 按省份过滤，例如 `北京`、`上海` |
| `--max-pages` | 50 | 最多爬取页数 |
| `--delay` | 1.0 | 每页请求间隔（秒） |
| `--output-dir` | 当前目录 | 结果文件保存路径 |

#### 方式二：从内置数据集生成 Excel

```bash
python generate_excel.py
```

### 输出文件

| 文件 | 说明 |
|------|------|
| `institutions.xlsx` | **主数据文件**，格式化 Excel，带筛选/冻结/超链接 |
| `institutions.csv` | CSV 格式，UTF-8 BOM 编码，可直接用 Excel 打开 |
| `institutions.json` | JSON 格式，便于程序处理 |

### 运行测试

```bash
python -m pytest test_scraper.py -v
```
