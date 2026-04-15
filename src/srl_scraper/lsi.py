"""LSI メディエンス WEB総合検査案内スクレイパー

構造:
  分野一覧:   https://data.medience.co.jp/guide/field-{01-13}.html
  リストページ: https://data.medience.co.jp/guide/list-{XXXX}.html
  詳細ページ:   https://data.medience.co.jp/guide/guide-{XXXXXXXX}.html

リストページに検査一覧テーブルがあり、詳細ページに JLAC10 (<p class="text-jlac10">) がある。
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
    classify_jlac10,
)

logger = logging.getLogger(__name__)

LSI_BASE = "https://data.medience.co.jp"
LSI_TOP = f"{LSI_BASE}/guide/"
LSI_LAST_UPDATE_FILE = "lsi_last_update_date.txt"


def fetch_lsi_update_date(session: requests.Session) -> str:
    """トップページから掲載日を取得"""
    resp = session.get(LSI_TOP, headers=HEADERS, timeout=REQUEST_TIMEOUT)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "lxml")
    span = soup.find("span", class_="-small")
    if not span:
        return ""
    text = span.get_text(strip=True)
    # "掲載内容は、2026 年 4 月 1 日時点の情報です。" → "2026/04/01"
    m = re.search(r"(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日", text)
    if m:
        return f"{m.group(1)}/{int(m.group(2)):02d}/{int(m.group(3)):02d}"
    return text


def check_lsi_update_needed(output_dir: Path, session: requests.Session) -> tuple[bool, str, str]:
    """LSIの更新有無を確認"""
    remote_date = fetch_lsi_update_date(session)
    date_file = output_dir / LSI_LAST_UPDATE_FILE
    local_date = ""
    if date_file.exists():
        local_date = date_file.read_text(encoding="utf-8").strip()
    return remote_date != local_date, remote_date, local_date


def save_lsi_last_update_date(output_dir: Path, date_str: str) -> None:
    (output_dir / LSI_LAST_UPDATE_FILE).write_text(date_str, encoding="utf-8")


def _fetch(url: str, session: requests.Session, cache_dir: Path | None, cache_hours: float) -> tuple[str, bool]:
    """HTML取得（キャッシュ対応）"""
    cached = _read_cache(cache_dir, url, cache_hours)
    if cached is not None:
        return cached, True

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            resp = session.get(url, headers=HEADERS, timeout=REQUEST_TIMEOUT)
            resp.raise_for_status()
            resp.encoding = resp.apparent_encoding or "utf-8"
            _write_cache(cache_dir, url, resp.text)
            return resp.text, False
        except requests.RequestException as e:
            if attempt == MAX_RETRIES:
                raise
            wait = RETRY_BACKOFF * attempt
            logger.warning("LSI リトライ %d/%d: %s (%.0f秒後)", attempt, MAX_RETRIES, e, wait)
            time.sleep(wait)
    raise RuntimeError("unreachable")


def discover_list_pages(session: requests.Session) -> list[dict]:
    """全分野ページからリストページURLを収集"""
    all_lists = []
    for i in range(1, 14):
        url = f"{LSI_BASE}/guide/field-{i:02d}.html"
        try:
            resp = session.get(url, headers=HEADERS, timeout=REQUEST_TIMEOUT)
            if resp.status_code != 200:
                continue
            soup = BeautifulSoup(resp.text, "lxml")
            for a in soup.find_all("a", href=re.compile(r"/guide/list-")):
                href = a.get("href", "")
                if not href.startswith("http"):
                    href = f"{LSI_BASE}{href}"
                all_lists.append({
                    "url": href,
                    "name": a.get_text(strip=True),
                    "field_id": i,
                })
        except requests.RequestException as e:
            logger.warning("LSI field-%02d 取得エラー: %s", i, e)
        time.sleep(REQUEST_INTERVAL)
    return all_lists


def scrape_list_page(
    url: str, session: requests.Session, cache_dir: Path | None, cache_hours: float
) -> list[dict]:
    """リストページからdetailリンクと基本データを取得"""
    html, from_cache = _fetch(url, session, cache_dir, cache_hours)
    if not from_cache:
        time.sleep(REQUEST_INTERVAL)

    soup = BeautifulSoup(html, "lxml")
    tables = soup.find_all("table")
    if not tables:
        return []

    items = []
    main_table = tables[0]
    rows = main_table.find_all("tr")

    for tr in rows:
        cells = tr.find_all(["th", "td"])
        if len(cells) < 6:
            continue

        # ヘッダー行をスキップ（th要素がある or "項目コード"テキスト）
        first_text = cells[0].get_text(strip=True)
        if cells[0].name == "th" or "項目コード" in first_text:
            continue

        # 詳細リンク
        link = tr.find("a", href=re.compile(r"guide-\d+"))
        detail_url = ""
        if link:
            href = link.get("href", "")
            if href.startswith("http"):
                detail_url = href
            else:
                # "guide-XXXX.html" or "./guide-XXXX.html" → 絶対パスに
                filename = re.search(r"(guide-[\d]+\.html)", href)
                if filename:
                    detail_url = f"{LSI_BASE}/guide/{filename.group(1)}"

        item_code = _clean_text(cells[0].get_text())

        # 検査項目名（h3から取得）
        h3 = cells[1].find("h3") if len(cells) > 1 else None
        item_name = ""
        if h3:
            # small/spanを除いた主テキスト
            for child in h3.children:
                if child.name in ("small", "span", "div"):
                    continue
                text = child.get_text(strip=True) if hasattr(child, "get_text") else str(child).strip()
                if text:
                    item_name = text
                    break
            if not item_name:
                item_name = _clean_text(h3.get_text())

        # 材料
        material_dl = cells[2].find("dl", class_="material-wrap") if len(cells) > 2 else None
        material = ""
        if material_dl:
            dt = material_dl.find("dt")
            material = dt.get_text(strip=True) if dt else ""

        # 検査方法
        method = _clean_text(cells[5].get_text()) if len(cells) > 5 else ""

        # 基準値
        ref_p = cells[6].find("p", class_="text-fiducial_point") if len(cells) > 6 else None
        unit_p = cells[6].find("p", class_="text-unit") if len(cells) > 6 else None
        reference = ""
        if ref_p:
            unit = unit_p.get_text(strip=True) if unit_p else ""
            ref_val = ref_p.get_text(strip=True)
            reference = f"{ref_val}({unit})" if unit else ref_val

        items.append({
            "item_code": item_code,
            "item_name": item_name,
            "material": material,
            "method": method,
            "reference_value": reference,
            "detail_url": detail_url,
        })

    return items


def fetch_jlac10(url: str, session: requests.Session, cache_dir: Path | None, cache_hours: float) -> tuple[str, str, bool]:
    """詳細ページから JLAC10 を取得

    Returns: (jlac10, jlac10_status, from_cache)
    """
    html, from_cache = _fetch(url, session, cache_dir, cache_hours)
    soup = BeautifulSoup(html, "lxml")
    p = soup.find("p", class_="text-jlac10")
    jlac10_raw = p.get_text(strip=True) if p else ""
    jlac10 = jlac10_raw.replace("-", "")
    jlac10_status = classify_jlac10(jlac10)
    return jlac10, jlac10_status, from_cache


def scrape_all(
    output_dir: Path | None = None,
    use_cache: bool = True,
    cache_max_age_hours: float = 24.0,
    check_update: bool = False,
) -> dict | None:
    """LSI全リストページをスクレイピング"""
    output_dir = output_dir or Path("data")
    output_dir.mkdir(parents=True, exist_ok=True)

    session = requests.Session()
    now = datetime.now(timezone.utc)

    remote_date = ""
    if check_update:
        needed, remote_date, local_date = check_lsi_update_needed(output_dir, session)
        if not needed:
            logger.info("LSI: 更新なし (掲載日: %s)。スキップします。", remote_date)
            return None
        logger.info("LSI: 更新を検出 (掲載日: %s → %s)。取得を開始します。", local_date or "(初回)", remote_date)
        use_cache = False
        time.sleep(REQUEST_INTERVAL)

    cache_dir = output_dir / ".cache" / "lsi" if use_cache else None
    if cache_dir:
        cache_dir.mkdir(parents=True, exist_ok=True)

    logger.info("LSI: 分野一覧からリストページを収集中...")
    list_pages = discover_list_pages(session)
    logger.info("LSI: %d リストページ検出", len(list_pages))

    # リストページから検査項目を収集
    all_items_raw: list[dict] = []
    for lp in list_pages:
        logger.info("LSI: [%s] リスト取得中...", lp["name"])
        items = scrape_list_page(lp["url"], session, cache_dir, cache_max_age_hours)
        for item in items:
            item["category_name"] = lp["name"]
        all_items_raw.extend(items)
        logger.info("  → %d件", len(items))

    # 詳細ページからJLAC10を取得（重複URL排除）
    detail_urls = {item["detail_url"] for item in all_items_raw if item["detail_url"]}
    logger.info("LSI: 詳細ページ %d件からJLAC10を取得開始", len(detail_urls))

    jlac10_map: dict[str, tuple[str, str]] = {}  # url -> (jlac10, jlac10_status)
    errors = []
    stats = {"fetched": 0, "cached": 0}

    for i, url in enumerate(sorted(detail_urls)):
        try:
            jlac10, jlac10_status, from_cache = fetch_jlac10(url, session, cache_dir, cache_max_age_hours)
            jlac10_map[url] = (jlac10, jlac10_status)
            if from_cache:
                stats["cached"] += 1
            else:
                stats["fetched"] += 1
                if i < len(detail_urls) - 1:
                    time.sleep(REQUEST_INTERVAL)
        except Exception as e:
            logger.error("LSI 詳細取得エラー: %s - %s", url, e)
            errors.append({"url": url, "error": str(e)})

        if (i + 1) % 50 == 0:
            logger.info("LSI: JLAC10取得 %d/%d 完了", i + 1, len(detail_urls))

    # JLAC10をマージ
    all_items = []
    for item in all_items_raw:
        jlac10_entry = jlac10_map.get(item["detail_url"])
        if jlac10_entry is None:
            jlac10, jlac10_status = "", "empty"
        else:
            jlac10, jlac10_status = jlac10_entry
        all_items.append({
            "jlac10": jlac10,
            "jlac10_status": jlac10_status,
            "item_name": item["item_name"],
            "material": item["material"],
            "method": item["method"],
            "reference_value": item["reference_value"],
            "lsi_code": item["item_code"],
            "category_name": item["category_name"],
            "detail_url": item["detail_url"],
            "source": "LSI",
        })

    result = {
        "metadata": {
            "source": "https://data.medience.co.jp/guide/",
            "scraped_at": now.isoformat(),
            "total_list_pages": len(list_pages),
            "total_items": len(all_items),
            "detail_pages_fetched": stats["fetched"],
            "detail_pages_cached": stats["cached"],
            "last_update_date": remote_date or None,
            "errors": errors,
        },
        "items": all_items,
    }

    timestamp = now.strftime("%Y%m%d_%H%M%S")
    filename = f"lsi_tests_{timestamp}.json"
    filepath = output_dir / filename
    filepath.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    logger.info("LSI 保存完了: %s (%d件)", filepath, len(all_items))

    latest = output_dir / "lsi_tests_latest.json"
    latest.unlink(missing_ok=True)
    latest.symlink_to(filename)

    if remote_date and not errors:
        save_lsi_last_update_date(output_dir, remote_date)

    return result
