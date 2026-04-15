"""SRL検査案内スクレイパー本体"""

import hashlib
import json
import logging
import re
import time
from datetime import datetime, timezone
from pathlib import Path

import requests
from bs4 import BeautifulSoup, Tag

from .categories import CATEGORIES, CATEGORY_BY_ID, Category

logger = logging.getLogger(__name__)

TOP_URL = "https://test-directory.srl.info/akiruno/"
BASE_URL = "https://test-directory.srl.info/akiruno/test/list/{}"
DETAIL_BASE = "https://test-directory.srl.info"
REQUEST_INTERVAL = 3.0  # 秒（サーバー負荷軽減）
REQUEST_TIMEOUT = 30  # 秒
MAX_RETRIES = 3
RETRY_BACKOFF = 5.0  # 秒（指数バックオフの基数）

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/120.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "ja,en;q=0.5",
}


# ---------------------------------------------------------------------------
# HTMLキャッシュ
# ---------------------------------------------------------------------------

def _cache_path(cache_dir: Path, url: str) -> Path:
    """URLからキャッシュファイルパスを生成"""
    h = hashlib.sha256(url.encode()).hexdigest()[:16]
    # URLの末尾の数字をファイル名に含めて可読性を確保
    m = re.search(r"/list/(\d+)$", url)
    suffix = f"_list_{m.group(1)}" if m else ""
    return cache_dir / f"{h}{suffix}.html"


def _read_cache(cache_dir: Path | None, url: str, max_age_hours: float) -> str | None:
    """キャッシュからHTMLを読み込む（有効期限内の場合のみ）"""
    if not cache_dir:
        return None
    path = _cache_path(cache_dir, url)
    if not path.exists():
        return None
    age_hours = (time.time() - path.stat().st_mtime) / 3600
    if age_hours > max_age_hours:
        logger.debug("キャッシュ期限切れ: %s (%.1f時間前)", path.name, age_hours)
        return None
    logger.info("  キャッシュ使用: %s (%.1f時間前)", path.name, age_hours)
    return path.read_text(encoding="utf-8")


def _write_cache(cache_dir: Path | None, url: str, html: str) -> None:
    """HTMLをキャッシュに保存"""
    if not cache_dir:
        return
    cache_dir.mkdir(parents=True, exist_ok=True)
    path = _cache_path(cache_dir, url)
    path.write_text(html, encoding="utf-8")


# ---------------------------------------------------------------------------
# HTMLパーサー
# ---------------------------------------------------------------------------

def _clean_text(text: str | None) -> str:
    """HTML要素からテキストを抽出し、空白を正規化する"""
    if not text:
        return ""
    return re.sub(r"\s+", " ", text.strip())


def _extract_storage(td: Tag) -> dict[str, str]:
    """保存条件セルからアイコンのalt属性と安定性期間を抽出"""
    img = td.find("img")
    condition = img.get("alt", "") if img else ""
    raw = _clean_text(td.get_text())
    stability = ""
    m = re.search(r"[（(](.+?)[）)]", raw)
    if m:
        stability = m.group(1)
    return {"condition": condition, "stability": stability}


def _extract_fee(td: Tag) -> dict[str, str | None]:
    """実施料・判断料セルをパース"""
    has_exclamation = td.find("img", src=re.compile(r"exclamation")) is not None
    raw = _clean_text(td.get_text())
    parts = [p.strip() for p in raw.split() if p.strip()]
    fee = parts[0] if parts else ""
    note = parts[1] if len(parts) > 1 else None
    return {
        "fee": fee,
        "note": note,
        "包括": has_exclamation,
    }


def _extract_method(td: Tag) -> dict[str, str]:
    """検査方法セルからメソッド名とヘルプテキストを抽出"""
    baloon = td.find("div", class_="baloon")
    help_text = ""
    if baloon:
        help_text = _clean_text(baloon.get_text())
        baloon.decompose()
    btn = td.find("input", class_="btn_help")
    if btn:
        btn.decompose()
    method = _clean_text(td.get_text())
    return {"name": method, "description": help_text}


def _extract_cap_color(td: Tag) -> str:
    """キャップカラー画像からコードを抽出"""
    img = td.find("img", class_="img_chap")
    if not img:
        return ""
    src = img.get("src", "")
    m = re.search(r"/([A-Z0-9]+)-c\.png", src)
    return m.group(1) if m else ""


def parse_test_row(tr: Tag) -> dict | None:
    """1行の<tr>から検査項目データを抽出"""
    link_url = tr.get("link_url", "")
    th = tr.find("th")
    if not th:
        return None

    p_tag = th.find("p")
    if not p_tag:
        return None
    lines = [_clean_text(line) for line in p_tag.stripped_strings]
    item_name = lines[0] if lines else ""
    jlac10_raw = lines[1] if len(lines) > 1 else ""
    jlac10 = jlac10_raw.replace("-", "")

    tds = tr.find_all("td")
    if len(tds) < 8:
        logger.warning("カラム数不足: %s (%d列)", item_name, len(tds))
        return None

    material_td = tds[0]
    material_lines = [_clean_text(s) for s in material_td.stripped_strings]
    material = material_lines[0] if material_lines else ""
    volume_ml = material_lines[1] if len(material_lines) > 1 else ""

    container = _clean_text(tds[1].get_text())
    cap_color = _extract_cap_color(tds[2])
    storage = _extract_storage(tds[3])
    turnaround = _clean_text(tds[4].get_text())
    fee = _extract_fee(tds[5])
    method = _extract_method(tds[6])
    reference = _clean_text(tds[7].get_text())

    detail_url = f"{DETAIL_BASE}{link_url}" if link_url else ""

    return {
        "item_name": item_name,
        "jlac10": jlac10,
        "detail_url": detail_url,
        "material": material,
        "volume_ml": volume_ml,
        "container": container,
        "cap_color": cap_color,
        "storage": storage,
        "turnaround_days": turnaround,
        "fee": fee,
        "method": method,
        "reference_value": reference,
    }


# ---------------------------------------------------------------------------
# Last Up Date チェック
# ---------------------------------------------------------------------------

LAST_UPDATE_FILE = "last_update_date.txt"


def fetch_last_update_date(session: requests.Session) -> str:
    """トップページから Last Up Date を取得"""
    resp = session.get(TOP_URL, headers=HEADERS, timeout=REQUEST_TIMEOUT)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "lxml")
    p = soup.find("p", class_="modify_date")
    if not p:
        raise RuntimeError("Last Up Date が見つかりません")
    span = p.find("span")
    if not span:
        raise RuntimeError("Last Up Date の日付が見つかりません")
    return span.get_text(strip=True)


def check_update_needed(output_dir: Path, session: requests.Session) -> tuple[bool, str, str]:
    """サイトの更新有無を確認する。
    Returns: (update_needed, remote_date, local_date)
    """
    remote_date = fetch_last_update_date(session)

    date_file = output_dir / LAST_UPDATE_FILE
    local_date = ""
    if date_file.exists():
        local_date = date_file.read_text(encoding="utf-8").strip()

    return remote_date != local_date, remote_date, local_date


def save_last_update_date(output_dir: Path, date_str: str) -> None:
    """取得成功後に日付を保存"""
    date_file = output_dir / LAST_UPDATE_FILE
    date_file.write_text(date_str, encoding="utf-8")


# ---------------------------------------------------------------------------
# フェッチ & スクレイピング
# ---------------------------------------------------------------------------

def fetch_page(
    url: str,
    session: requests.Session,
    cache_dir: Path | None = None,
    cache_max_age_hours: float = 24.0,
) -> tuple[str, bool]:
    """HTMLを取得する。キャッシュがあればそちらを使用。
    Returns: (html, from_cache)
    """
    cached = _read_cache(cache_dir, url, cache_max_age_hours)
    if cached is not None:
        return cached, True

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            resp = session.get(url, headers=HEADERS, timeout=REQUEST_TIMEOUT)
            resp.raise_for_status()
            resp.encoding = resp.apparent_encoding
            html = resp.text
            _write_cache(cache_dir, url, html)
            return html, False
        except requests.RequestException as e:
            if attempt == MAX_RETRIES:
                raise
            wait = RETRY_BACKOFF * attempt
            logger.warning("リトライ %d/%d: %s (%.0f秒後)", attempt, MAX_RETRIES, e, wait)
            time.sleep(wait)
    raise RuntimeError("unreachable")


def scrape_category(
    category: Category,
    session: requests.Session,
    cache_dir: Path | None = None,
    cache_max_age_hours: float = 24.0,
) -> tuple[list[dict], bool]:
    """1カテゴリページをスクレイピング。
    Returns: (items, from_cache)
    """
    url = BASE_URL.format(category.id)
    logger.info("取得中: [%d] %s - %s", category.id, category.group, category.name)

    html, from_cache = fetch_page(url, session, cache_dir, cache_max_age_hours)
    soup = BeautifulSoup(html, "lxml")

    container = soup.find("div", class_="list-container-div")
    if not container:
        logger.warning("テーブルコンテナが見つかりません: %s", url)
        return [], from_cache

    table = container.find("table", class_="list_culomn9")
    if not table:
        logger.warning("テーブルが見つかりません: %s", url)
        return [], from_cache

    # サーバーHTMLには<tbody>がない場合がある（ブラウザが自動挿入する）
    tbody = table.find("tbody")
    search_root = tbody if tbody else table

    rows = search_root.find_all("tr", class_="with_link")
    items = []
    for tr in rows:
        item = parse_test_row(tr)
        if item:
            item["category_id"] = category.id
            item["category_name"] = category.name
            item["category_group"] = category.group
            items.append(item)

    logger.info("  → %d件の検査項目を取得", len(items))
    return items, from_cache


def scrape_all(
    category_ids: list[int] | None = None,
    output_dir: Path | None = None,
    use_cache: bool = True,
    cache_max_age_hours: float = 24.0,
    check_update: bool = False,
    force: bool = False,
) -> dict | None:
    """全カテゴリ（または指定カテゴリ）をスクレイピングしてJSONで保存

    Args:
        category_ids: 取得対象のカテゴリID一覧。Noneなら全カテゴリ。
        output_dir: JSON出力先ディレクトリ。
        use_cache: HTMLキャッシュを使用するか。
        cache_max_age_hours: キャッシュの有効期限（時間）。
        check_update: Trueならトップページの Last Up Date を確認し、
                      更新がなければスキップする。
        force: check_update=True でも強制的にスクレイピングする。

    Returns:
        取得結果のdict。更新不要でスキップした場合はNone。
    """
    output_dir = output_dir or Path("data")
    output_dir.mkdir(parents=True, exist_ok=True)

    session = requests.Session()

    # Last Up Date チェック
    remote_date = ""
    if check_update:
        needed, remote_date, local_date = check_update_needed(output_dir, session)
        if not needed and not force:
            logger.info(
                "更新なし (Last Up Date: %s → %s)。スキップします。",
                local_date, remote_date,
            )
            return None
        if needed:
            logger.info(
                "更新を検出 (Last Up Date: %s → %s)。取得を開始します。",
                local_date or "(初回)", remote_date,
            )
            # 日付が変わったのでキャッシュを無効化
            use_cache = False
        time.sleep(REQUEST_INTERVAL)

    cache_dir = output_dir / ".cache" if use_cache else None

    if category_ids:
        targets = [CATEGORY_BY_ID[cid] for cid in category_ids if cid in CATEGORY_BY_ID]
        if not targets:
            raise ValueError(f"有効なカテゴリIDがありません: {category_ids}")
    else:
        targets = list(CATEGORIES)

    all_items: list[dict] = []
    errors: list[dict] = []
    stats = {"fetched": 0, "cached": 0}
    now = datetime.now(timezone.utc)

    for i, cat in enumerate(targets):
        try:
            items, from_cache = scrape_category(
                cat, session, cache_dir, cache_max_age_hours
            )
            all_items.extend(items)
            if from_cache:
                stats["cached"] += 1
            else:
                stats["fetched"] += 1
        except Exception as e:
            logger.error("カテゴリ %d (%s) でエラー: %s", cat.id, cat.name, e)
            errors.append({"category_id": cat.id, "name": cat.name, "error": str(e)})

        # レート制限：サーバーからフェッチした場合のみ待機
        if i < len(targets) - 1 and not (use_cache and stats["cached"] == i + 1):
            time.sleep(REQUEST_INTERVAL)

    result = {
        "metadata": {
            "source": "https://test-directory.srl.info/akiruno",
            "scraped_at": now.isoformat(),
            "last_update_date": remote_date or None,
            "total_categories": len(targets),
            "total_items": len(all_items),
            "fetched_from_server": stats["fetched"],
            "served_from_cache": stats["cached"],
            "errors": errors,
        },
        "items": all_items,
    }

    # ファイル出力
    timestamp = now.strftime("%Y%m%d_%H%M%S")
    filename = f"srl_tests_{timestamp}.json"
    filepath = output_dir / filename
    filepath.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    logger.info("保存完了: %s (%d件)", filepath, len(all_items))

    # latest シンボリックリンクを更新
    latest = output_dir / "srl_tests_latest.json"
    latest.unlink(missing_ok=True)
    latest.symlink_to(filename)
    logger.info("最新リンク更新: %s → %s", latest, filename)

    # Last Up Date を記録（次回のチェック用）
    if remote_date and not errors:
        save_last_update_date(output_dir, remote_date)
        logger.info("Last Up Date 記録: %s", remote_date)

    return result


def diff_report(old_path: Path, new_path: Path) -> dict:
    """2つのJSONファイルを比較して差分レポートを生成"""
    old_data = json.loads(old_path.read_text(encoding="utf-8"))
    new_data = json.loads(new_path.read_text(encoding="utf-8"))

    def _key(item: dict) -> str:
        return f"{item['category_id']}:{item['jlac10']}:{item['detail_url']}"

    old_items = {_key(i): i for i in old_data["items"]}
    new_items = {_key(i): i for i in new_data["items"]}

    old_keys = set(old_items.keys())
    new_keys = set(new_items.keys())

    added = [new_items[k] for k in sorted(new_keys - old_keys)]
    removed = [old_items[k] for k in sorted(old_keys - new_keys)]

    changed = []
    for k in sorted(old_keys & new_keys):
        if old_items[k] != new_items[k]:
            changed.append({
                "key": k,
                "old": old_items[k],
                "new": new_items[k],
            })

    report = {
        "compared": {
            "old": str(old_path),
            "new": str(new_path),
            "generated_at": datetime.now(timezone.utc).isoformat(),
        },
        "summary": {
            "added": len(added),
            "removed": len(removed),
            "changed": len(changed),
            "unchanged": len(old_keys & new_keys) - len(changed),
        },
        "added": added,
        "removed": removed,
        "changed": changed,
    }
    return report
