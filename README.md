# zsh_code_test

## 研招网院校数据爬取工具

使用 Python 从[中国研究生招生信息网（研招网）](https://yz.chsi.com.cn)拉取招生院校数据，结果保存为 CSV 和 JSON 文件。

### 功能

- 爬取全国或指定省份的招生院校列表
- 解析院校名称、所在省份、院校类别及详情链接
- 自动翻页，支持最大页数限制
- 保存为 `institutions.csv`（Excel 友好的 UTF-8 BOM 编码）和 `institutions.json`

### 环境要求

- Python 3.10+

### 安装依赖

```bash
pip install -r requirements.txt
```

### 使用方法

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

### 输出文件

| 文件 | 说明 |
|------|------|
| `institutions.csv` | 院校数据，UTF-8 BOM 编码，可直接用 Excel 打开 |
| `institutions.json` | 院校数据，JSON 格式，便于程序处理 |

字段：`name`（院校名称）、`province`（省份）、`category`（院校类别）、`url`（详情链接）

### 运行测试

```bash
python -m pytest test_scraper.py -v
```
