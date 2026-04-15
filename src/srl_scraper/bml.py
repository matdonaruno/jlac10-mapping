"""BML検査案内スクレイパー

構造:
  カテゴリ一覧: https://uwb01.bml.co.jp/kensa/search/
  カテゴリ結果: https://uwb01.bml.co.jp/kensa/search/result/{group}/{sub}
  検査詳細:     https://uwb01.bml.co.jp/kensa/search/detail/{code}

詳細ページに「統一コード」(= JLAC10) がある。
"""

import json
import logging
import re
import time
from datetime import datetime, timezone
from pathlib import Path

import requests
from bs4 import BeautifulSoup

from .scraper import (
    HEADERS,
    MAX_RETRIES,
    REQUEST_INTERVAL,
    REQUEST_TIMEOUT,
    RETRY_BACKOFF,
    _clean_text,
    _read_cache,
    _write_cache,
)

logger = logging.getLogger(__name__)

BML_BASE = "https://uwb01.bml.co.jp"
BML_SEARCH = f"{BML_BASE}/kensa/search/"
BML_NEW = f"{BML_BASE}/kensa/new"
BML_CHANGE = f"{BML_BASE}/kensa/change"
BML_LAST_UPDATE_FILE = "bml_last_update_date.txt"


def _fetch(url: str, session: requests.Session, cache_dir: Path | None, cache_hours: float) -> tuple[str, bool]:
    """HTML取得（キャッシュ対応）"""
    cached = _read_cache(cache_dir, url, cache_hours)
    if cached is not None:
        return cached, True

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            resp = session.get(url, headers=HEADERS, timeout=REQUEST_TIMEOUT)
            resp.raise_for_status()
            resp.encoding = resp.apparent_encoding
            _write_cache(cache_dir, url, resp.text)
            return resp.text, False
        except requests.RequestException as e:
            if attempt == MAX_RETRIES:
                raise
            wait = RETRY_BACKOFF * attempt
            logger.warning("BML リトライ %d/%d: %s (%.0f秒後)", attempt, MAX_RETRIES, e, wait)
            time.sleep(wait)
    raise RuntimeError("unreachable")


def _get_latest_date_from_table(url: str, session: requests.Session) -> str:
    """new/change ページのテーブルから最新の掲載日を取得"""
    resp = session.get(url, headers=HEADERS, timeout=REQUEST_TIMEOUT)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "lxml")
    tables = soup.find_all("table")
    if not tables:
        return ""
    # 最初のデータ行の「掲載日」列（2列目）
    for tr in tables[0].find_all("tr")[1:2]:
        cells = tr.find_all("td")
        if len(cells) >= 2:
            return _clean_text(cells[1].get_text())  # 掲載日
    return ""


def fetch_bml_latest_date(session: requests.Session) -> str:
    """new と change の掲載日で新しい方を返す"""
    new_date = _get_latest_date_from_table(BML_NEW, session)
    time.sleep(REQUEST_INTERVAL)
    change_date = _get_latest_date_from_table(BML_CHANGE, session)
    # 日付文字列として比較（YYYY/MM/DD形式）
    return max(new_date, change_date) if new_date or change_date else ""


def check_bml_update_needed(output_dir: Path, session: requests.Session) -> tuple[bool, str, str]:
    """BMLの更新有無を確認"""
    remote_date = fetch_bml_latest_date(session)
    date_file = output_dir / BML_LAST_UPDATE_FILE
    local_date = ""
    if date_file.exists():
        local_date = date_file.read_text(encoding="utf-8").strip()
    return remote_date != local_date, remote_date, local_date


def save_bml_last_update_date(output_dir: Path, date_str: str) -> None:
    (output_dir / BML_LAST_UPDATE_FILE).write_text(date_str, encoding="utf-8")


def discover_categories(session: requests.Session) -> list[dict]:
    """検索トップページからカテゴリ一覧を取得"""
    resp = session.get(BML_SEARCH, headers=HEADERS, timeout=REQUEST_TIMEOUT)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "lxml")

    categories = []
    for a in soup.find_all("a", href=re.compile(r"/kensa/search/result/")):
        text = a.get_text(strip=True)
        if text == "すべて":
            continue
        href = a.get("href", "")
        categories.append({"url": f"{BML_BASE}{href}", "name": text})
    return categories


def scrape_category_list(url: str, session: requests.Session, cache_dir: Path | None, cache_hours: float) -> list[str]:
    """カテゴリ結果ページから詳細ページURLを取得"""
    html, from_cache = _fetch(url, session, cache_dir, cache_hours)
    if not from_cache:
        time.sleep(REQUEST_INTERVAL)
    soup = BeautifulSoup(html, "lxml")
    links = soup.find_all("a", href=re.compile(r"/kensa/search/detail/"))
    return [f"{BML_BASE}{a.get('href')}" for a in links]


def parse_detail(html: str, url: str) -> dict | None:
    """BML詳細ページから検査データを抽出"""
    soup = BeautifulSoup(html, "lxml")
    tables = soup.find_all("table")
    if not tables:
        return None

    main_table = tables[0]
    rows = main_table.find_all("tr")

    data = {}
    for tr in rows:
        cells = tr.find_all(["th", "td"])
        if len(cells) < 2:
            continue
        key = _clean_text(cells[0].get_text())
        val = _clean_text(cells[1].get_text())
        data[key] = val

    jlac10_raw = data.get("統一コード", "")
    jlac10 = jlac10_raw.replace("-", "")
    if not jlac10:
        return None

    # 検体必要量から材料を抽出
    material_raw = data.get("検体必要量(mL)容器 / 保存", "")
    # "血清 0.5B-1S-1(1か月)" のような形式
    material = material_raw.split()[0] if material_raw else ""

    return {
        "jlac10": jlac10,
        "item_name": data.get("検査項目名称", ""),
        "material": material,
        "method": data.get("検査方法", ""),
        "reference_value": data.get("基準値", ""),
        "bml_code": data.get("コード", ""),
        "detail_url": url,
        "source": "BML",
    }


def scrape_all(
    output_dir: Path | None = None,
    use_cache: bool = True,
    cache_max_age_hours: float = 24.0,
    check_update: bool = False,
) -> dict | None:
    """BML全カテゴリをスクレイピング"""
    output_dir = output_dir or Path("data")
    output_dir.mkdir(parents=True, exist_ok=True)

    session = requests.Session()
    now = datetime.now(timezone.utc)

    remote_date = ""
    if check_update:
        needed, remote_date, local_date = check_bml_update_needed(output_dir, session)
        if not needed:
            logger.info("BML: 更新なし (最新掲載日: %s)。スキップします。", remote_date)
            return None
        logger.info("BML: 更新を検出 (掲載日: %s → %s)。取得を開始します。", local_date or "(初回)", remote_date)
        use_cache = False
        time.sleep(REQUEST_INTERVAL)

    cache_dir = output_dir / ".cache" / "bml" if use_cache else None
    if cache_dir:
        cache_dir.mkdir(parents=True, exist_ok=True)

    logger.info("BML: カテゴリ一覧を取得中...")
    categories = discover_categories(session)
    logger.info("BML: %d カテゴリ検出", len(categories))
    time.sleep(REQUEST_INTERVAL)

    # 全カテゴリから詳細URLを収集（重複除去）
    detail_urls: dict[str, str] = {}  # url -> category_name
    for cat in categories:
        logger.info("BML: カテゴリ [%s] のリスト取得中...", cat["name"])
        urls = scrape_category_list(cat["url"], session, cache_dir, cache_max_age_hours)
        for u in urls:
            if u not in detail_urls:
                detail_urls[u] = cat["name"]
        logger.info("  → %d件", len(urls))

    logger.info("BML: 詳細ページ %d件を取得開始", len(detail_urls))

    all_items = []
    errors = []
    stats = {"fetched": 0, "cached": 0}

    for i, (url, cat_name) in enumerate(detail_urls.items()):
        try:
            html, from_cache = _fetch(url, session, cache_dir, cache_max_age_hours)
            if from_cache:
                stats["cached"] += 1
            else:
                stats["fetched"] += 1
                if i < len(detail_urls) - 1:
                    time.sleep(REQUEST_INTERVAL)

            item = parse_detail(html, url)
            if item:
                item["category_name"] = cat_name
                all_items.append(item)
        except Exception as e:
            logger.error("BML 詳細取得エラー: %s - %s", url, e)
            errors.append({"url": url, "error": str(e)})

        if (i + 1) % 50 == 0:
            logger.info("BML: %d/%d 完了", i + 1, len(detail_urls))

    result = {
        "metadata": {
            "source": "https://uwb01.bml.co.jp/kensa/search/",
            "scraped_at": now.isoformat(),
            "total_categories": len(categories),
            "total_items": len(all_items),
            "fetched_from_server": stats["fetched"],
            "served_from_cache": stats["cached"],
            "last_update_date": remote_date or None,
            "errors": errors,
        },
        "items": all_items,
    }

    timestamp = now.strftime("%Y%m%d_%H%M%S")
    filename = f"bml_tests_{timestamp}.json"
    filepath = output_dir / filename
    filepath.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    logger.info("BML 保存完了: %s (%d件)", filepath, len(all_items))

    latest = output_dir / "bml_tests_latest.json"
    latest.unlink(missing_ok=True)
    latest.symlink_to(filename)

    if remote_date and not errors:
        save_bml_last_update_date(output_dir, remote_date)

    return result
