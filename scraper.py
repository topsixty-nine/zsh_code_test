"""
中国研究生招生信息网（研招网）院校数据爬取脚本
网站: https://yz.chsi.com.cn

功能:
  - 从研招网搜索并拉取招生院校列表
  - 支持按省份/学科门类过滤
  - 将结果保存为 CSV 和 JSON 文件
"""

import csv
import json
import logging
import time
from pathlib import Path
from urllib.parse import urljoin, urlparse

import requests
from bs4 import BeautifulSoup

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
logger = logging.getLogger(__name__)

BASE_URL = "https://yz.chsi.com.cn"
SEARCH_URL = f"{BASE_URL}/zsb/zsbss/search.do"

# 省份代码映射（部分常用省份）
PROVINCE_CODES = {
    "北京": "11",
    "天津": "12",
    "河北": "13",
    "山西": "14",
    "内蒙古": "15",
    "辽宁": "21",
    "吉林": "22",
    "黑龙江": "23",
    "上海": "31",
    "江苏": "32",
    "浙江": "33",
    "安徽": "34",
    "福建": "35",
    "江西": "36",
    "山东": "37",
    "河南": "41",
    "湖北": "42",
    "湖南": "43",
    "广东": "44",
    "广西": "45",
    "海南": "46",
    "重庆": "50",
    "四川": "51",
    "贵州": "52",
    "云南": "53",
    "西藏": "54",
    "陕西": "61",
    "甘肃": "62",
    "青海": "63",
    "宁夏": "64",
    "新疆": "65",
}

DEFAULT_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Referer": BASE_URL,
    "Accept-Language": "zh-CN,zh;q=0.9",
}


def fetch_page(session: requests.Session, page: int, province_code: str = "") -> str | None:
    """
    请求研招网院校搜索页，返回 HTML 文本。

    :param session: requests.Session 对象（保持 Cookie）
    :param page: 页码（从 1 开始）
    :param province_code: 省份代码，为空则不过滤
    :return: HTML 字符串，请求失败时返回 None
    """
    params = {
        "pg": page,
        "provinceId": province_code,
        "yxmc": "",
    }
    try:
        response = session.get(SEARCH_URL, params=params, timeout=15)
        response.raise_for_status()
        return response.text
    except requests.RequestException as exc:
        logger.error("请求第 %d 页时出错: %s", page, exc)
        return None


def parse_institutions(html: str) -> list[dict]:
    """
    从 HTML 页面中解析院校信息列表。

    :param html: 研招网院校搜索结果页面 HTML
    :return: 院校信息字典列表，每条包含 name / province / category / url 字段
    """
    soup = BeautifulSoup(html, "lxml")
    institutions: list[dict] = []

    # 院校列表通常在 class="search-list" 或 "yxxx-list" 下
    rows = soup.select("ul.search-list li") or soup.select("div.yxxx-list li")

    for row in rows:
        name_tag = row.select_one("a.yxmc") or row.select_one("a")
        province_tag = row.select_one("span.yxsf") or row.select_one("span.province")
        category_tag = row.select_one("span.yxlb") or row.select_one("span.category")

        if not name_tag:
            continue

        href = name_tag.get("href", "")
        full_url = urljoin(BASE_URL, href) if href else BASE_URL

        institutions.append(
            {
                "name": name_tag.get_text(strip=True),
                "province": province_tag.get_text(strip=True) if province_tag else "",
                "category": category_tag.get_text(strip=True) if category_tag else "",
                "url": full_url,
            }
        )

    return institutions


def get_total_pages(html: str) -> int:
    """
    从页面中提取总页数。

    :param html: 研招网院校搜索结果页面 HTML
    :return: 总页数，无法解析时返回 1
    """
    soup = BeautifulSoup(html, "lxml")

    # 尝试从分页控件中获取最后一页的页码
    page_tags = soup.select("div.pagination a") or soup.select("div.pages a")
    pages = []
    for tag in page_tags:
        text = tag.get_text(strip=True)
        if text.isdigit():
            pages.append(int(text))

    # 有些页面在 input 或 span 里标注总页数
    if not pages:
        total_tag = soup.select_one("span.total-page") or soup.select_one("input[name='totalPage']")
        if total_tag:
            val = total_tag.get("value") or total_tag.get_text(strip=True)
            if val and val.isdigit():
                return int(val)

    return max(pages, default=1)


def scrape(
    province: str = "",
    max_pages: int = 50,
    delay: float = 1.0,
) -> list[dict]:
    """
    爬取研招网院校数据。

    :param province: 省份中文名（如"北京"），为空则爬取全国
    :param max_pages: 最多爬取页数，防止无限翻页
    :param delay: 每页请求间隔（秒），降低对服务器的压力
    :return: 所有院校信息列表
    """
    province_code = PROVINCE_CODES.get(province, "")
    if province and not province_code:
        logger.warning("未找到省份 '%s' 的代码，将爬取全国数据", province)

    session = requests.Session()
    session.headers.update(DEFAULT_HEADERS)

    all_institutions: list[dict] = []

    logger.info("开始爬取第 1 页…")
    first_html = fetch_page(session, 1, province_code)
    if not first_html:
        logger.error("无法获取第一页，终止")
        return all_institutions

    total_pages = min(get_total_pages(first_html), max_pages)
    logger.info("共 %d 页（上限 %d）", total_pages, max_pages)

    institutions = parse_institutions(first_html)
    all_institutions.extend(institutions)
    logger.info("第 1 页解析到 %d 所院校", len(institutions))

    for page in range(2, total_pages + 1):
        time.sleep(delay)
        logger.info("正在爬取第 %d/%d 页…", page, total_pages)
        html = fetch_page(session, page, province_code)
        if not html:
            continue
        institutions = parse_institutions(html)
        all_institutions.extend(institutions)
        logger.info("第 %d 页解析到 %d 所院校", page, len(institutions))

    logger.info("共爬取到 %d 所院校", len(all_institutions))
    return all_institutions


def save_csv(data: list[dict], filepath: str | Path) -> None:
    """将院校数据保存为 CSV 文件。"""
    filepath = Path(filepath)
    if not data:
        logger.warning("数据为空，跳过 CSV 保存")
        return

    fieldnames = list(data[0].keys())
    with filepath.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(data)
    logger.info("已保存 CSV: %s", filepath)


def save_json(data: list[dict], filepath: str | Path) -> None:
    """将院校数据保存为 JSON 文件。"""
    filepath = Path(filepath)
    with filepath.open("w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    logger.info("已保存 JSON: %s", filepath)


def main() -> None:
    """命令行入口：爬取全国院校数据并保存到当前目录。"""
    import argparse

    parser = argparse.ArgumentParser(
        description="从中国研究生招生信息网（研招网）拉取院校数据"
    )
    parser.add_argument(
        "--province",
        default="",
        help="按省份过滤，例如 '北京'（默认：全国）",
    )
    parser.add_argument(
        "--max-pages",
        type=int,
        default=50,
        help="最多爬取页数（默认：50）",
    )
    parser.add_argument(
        "--delay",
        type=float,
        default=1.0,
        help="每页请求间隔秒数（默认：1.0）",
    )
    parser.add_argument(
        "--output-dir",
        default=".",
        help="结果文件保存目录（默认：当前目录）",
    )
    args = parser.parse_args()

    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    data = scrape(
        province=args.province,
        max_pages=args.max_pages,
        delay=args.delay,
    )

    if data:
        save_csv(data, output_dir / "institutions.csv")
        save_json(data, output_dir / "institutions.json")
    else:
        logger.warning("未获取到任何数据，请检查网络或网站结构是否有变化")


if __name__ == "__main__":
    main()
