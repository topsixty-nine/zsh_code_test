"""
Unit tests for scraper.py

Tests cover HTML parsing helpers and CSV/JSON save utilities using
mock HTTP responses so no real network calls are made.
"""

import json
import textwrap
from pathlib import Path
from unittest.mock import MagicMock, patch
from urllib.parse import urlparse

import pytest

from scraper import (
    get_total_pages,
    parse_institutions,
    save_csv,
    save_json,
    scrape,
)

# ---------------------------------------------------------------------------
# Fixtures / helpers
# ---------------------------------------------------------------------------

SAMPLE_HTML = textwrap.dedent(
    """\
    <html><body>
      <ul class="search-list">
        <li>
          <a class="yxmc" href="/zsb/zsbxx.do?id=1001">北京大学</a>
          <span class="yxsf">北京</span>
          <span class="yxlb">综合</span>
        </li>
        <li>
          <a class="yxmc" href="/zsb/zsbxx.do?id=1002">清华大学</a>
          <span class="yxsf">北京</span>
          <span class="yxlb">理工</span>
        </li>
      </ul>
      <div class="pagination">
        <a href="?pg=1">1</a>
        <a href="?pg=2">2</a>
        <a href="?pg=3">3</a>
      </div>
    </body></html>
    """
)

EMPTY_HTML = "<html><body><ul class='search-list'></ul></body></html>"


# ---------------------------------------------------------------------------
# parse_institutions
# ---------------------------------------------------------------------------

def test_parse_institutions_returns_correct_count():
    result = parse_institutions(SAMPLE_HTML)
    assert len(result) == 2


def test_parse_institutions_fields():
    result = parse_institutions(SAMPLE_HTML)
    first = result[0]
    assert first["name"] == "北京大学"
    assert first["province"] == "北京"
    assert first["category"] == "综合"
    assert "1001" in first["url"]


def test_parse_institutions_full_url_for_relative_href():
    result = parse_institutions(SAMPLE_HTML)
    parsed = urlparse(result[0]["url"])
    assert parsed.scheme == "https"
    assert parsed.netloc == "yz.chsi.com.cn"


def test_parse_institutions_empty_html():
    result = parse_institutions(EMPTY_HTML)
    assert result == []


# ---------------------------------------------------------------------------
# get_total_pages
# ---------------------------------------------------------------------------

def test_get_total_pages_reads_pagination():
    assert get_total_pages(SAMPLE_HTML) == 3


def test_get_total_pages_defaults_to_one_when_no_pagination():
    assert get_total_pages(EMPTY_HTML) == 1


def test_get_total_pages_reads_span_total():
    html = "<html><body><span class='total-page'>7</span></body></html>"
    assert get_total_pages(html) == 7


# ---------------------------------------------------------------------------
# save_csv / save_json
# ---------------------------------------------------------------------------

def test_save_csv_creates_file(tmp_path):
    data = [{"name": "北京大学", "province": "北京", "category": "综合", "url": "http://example.com"}]
    out = tmp_path / "out.csv"
    save_csv(data, out)
    assert out.exists()
    content = out.read_text(encoding="utf-8-sig")
    assert "北京大学" in content


def test_save_csv_header_row(tmp_path):
    data = [{"name": "清华大学", "province": "北京", "category": "理工", "url": "http://example.com"}]
    out = tmp_path / "out.csv"
    save_csv(data, out)
    lines = out.read_text(encoding="utf-8-sig").splitlines()
    assert lines[0] == "name,province,category,url"


def test_save_csv_empty_data_no_file(tmp_path, caplog):
    out = tmp_path / "out.csv"
    save_csv([], out)
    assert not out.exists()
    assert "数据为空" in caplog.text


def test_save_json_creates_file(tmp_path):
    data = [{"name": "复旦大学", "province": "上海", "category": "综合", "url": "http://example.com"}]
    out = tmp_path / "out.json"
    save_json(data, out)
    assert out.exists()
    loaded = json.loads(out.read_text(encoding="utf-8"))
    assert loaded[0]["name"] == "复旦大学"


def test_save_json_valid_json_structure(tmp_path):
    data = [
        {"name": "浙江大学", "province": "浙江", "category": "综合", "url": "http://a.com"},
        {"name": "上海交通大学", "province": "上海", "category": "理工", "url": "http://b.com"},
    ]
    out = tmp_path / "out.json"
    save_json(data, out)
    loaded = json.loads(out.read_text(encoding="utf-8"))
    assert len(loaded) == 2


# ---------------------------------------------------------------------------
# scrape (mocked network)
# ---------------------------------------------------------------------------

def _mock_session_get(html: str):
    """Return a factory that makes requests.Session.get return a mock response."""
    mock_resp = MagicMock()
    mock_resp.raise_for_status = MagicMock()
    mock_resp.text = html
    return mock_resp


@patch("scraper.requests.Session")
def test_scrape_single_page(mock_session_cls):
    session = MagicMock()
    session.get.return_value = _mock_session_get(SAMPLE_HTML)
    mock_session_cls.return_value = session

    result = scrape(max_pages=1, delay=0)

    assert len(result) == 2
    assert result[0]["name"] == "北京大学"


@patch("scraper.requests.Session")
def test_scrape_returns_empty_on_network_failure(mock_session_cls):
    import requests as req_lib

    session = MagicMock()
    session.get.side_effect = req_lib.RequestException("connection error")
    mock_session_cls.return_value = session

    result = scrape(max_pages=1, delay=0)
    assert result == []


@patch("scraper.requests.Session")
def test_scrape_province_filter(mock_session_cls):
    session = MagicMock()
    session.get.return_value = _mock_session_get(SAMPLE_HTML)
    mock_session_cls.return_value = session

    scrape(province="北京", max_pages=1, delay=0)

    call_kwargs = session.get.call_args
    params = call_kwargs[1].get("params") or call_kwargs[0][1]
    assert params.get("provinceId") == "11"


@patch("scraper.requests.Session")
def test_scrape_unknown_province_uses_empty_code(mock_session_cls):
    session = MagicMock()
    session.get.return_value = _mock_session_get(SAMPLE_HTML)
    mock_session_cls.return_value = session

    scrape(province="火星", max_pages=1, delay=0)

    call_kwargs = session.get.call_args
    params = call_kwargs[1].get("params") or call_kwargs[0][1]
    assert params.get("provinceId") == ""
