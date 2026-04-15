"""試薬マスター DB 構築

1. メーカー製品ページから試薬名 + 添付文書リンクを収集
2. PMDA HTMLページから使用目的・測定原理を抽出
3. 試薬 → 測定法マッピング JSON を生成

対応メーカー（順次追加）:
  - カイノス (https://www.kainos.co.jp/products/biochem/)
"""

import json
import logging
import re
import time
from datetime import datetime, timezone
from pathlib import Path

import requests
from bs4 import BeautifulSoup

from .scraper import HEADERS, REQUEST_INTERVAL, REQUEST_TIMEOUT

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# PMDA 添付文書 HTML パーサー
# ---------------------------------------------------------------------------

def parse_pmda_detail(html: str) -> dict:
    """PMDA 体外診断薬 添付文書 HTML から情報を抽出"""
    soup = BeautifulSoup(html, "lxml")

    def _section_text(heading_text: str, stop_at: str | None = None) -> str:
        """見出し直下の div.dd テキストのみ収集（次の h3/h4 まで）"""
        for h in soup.find_all(["h3", "h4"]):
            if heading_text not in h.get_text():
                continue
            texts = []
            for sib in h.find_next_siblings():
                if sib.name in ("h3", "h4", "hr"):
                    break
                if stop_at and stop_at in sib.get_text():
                    break
                # div.dd のテキストのみ
                if sib.name == "div" and "dd" in (sib.get("class") or []):
                    texts.append(sib.get_text(strip=True))
                # dl/dt/dd 内のテキスト
                elif sib.name == "dl":
                    texts.append(sib.get_text(strip=True))
            return "\n".join(texts)
        return ""

    # 販売名
    product_name = ""
    for h in soup.find_all("h4"):
        if "販売名" in h.get_text():
            dd = h.find_next("div", class_="dd")
            if dd:
                product_name = dd.get_text(strip=True)
            break

    # 一般的名称コード + 名称
    general_code = ""
    general_name = ""
    for h in soup.find_all("dt"):
        if "一般的名称" in h.get_text():
            dt_code = h.find_next("dt")
            if dt_code:
                code_text = dt_code.get_text(strip=True)
                if re.match(r"^\d+$", code_text):
                    general_code = code_text
                    dd = dt_code.find_next("div", class_="dd")
                    if dd:
                        general_name = dd.get_text(strip=True)
            break

    # 使用目的
    purpose = _section_text("使用目的")

    # 測定原理
    principle = _section_text("測定原理")

    # 製造販売業者
    manufacturer = ""
    for h in soup.find_all("h3"):
        if "製造販売業者" in h.get_text():
            dd = h.find_next("div", class_="dd")
            if dd:
                manufacturer = dd.get_text(strip=True)
            break

    return {
        "product_name": product_name,
        "general_code": general_code,
        "general_name": general_name,
        "purpose": purpose,
        "principle": principle,
        "manufacturer": manufacturer,
    }


def fetch_pmda_detail(url: str, session: requests.Session) -> dict:
    """PMDA 添付文書ページを取得してパース"""
    resp = session.get(url, headers=HEADERS, timeout=REQUEST_TIMEOUT)
    resp.raise_for_status()
    resp.encoding = "utf-8"
    return parse_pmda_detail(resp.text)


# ---------------------------------------------------------------------------
# メーカー製品ページ スクレイパー
# ---------------------------------------------------------------------------

def scrape_kainos_biochem(session: requests.Session) -> list[dict]:
    """カイノス 生化学検査試薬ページから試薬一覧を取得"""
    url = "https://www.kainos.co.jp/products/biochem/"
    resp = session.get(url, headers=HEADERS, timeout=REQUEST_TIMEOUT)
    resp.encoding = "utf-8"
    soup = BeautifulSoup(resp.text, "lxml")

    reagents = []
    for section in soup.find_all("section", class_="products-sec"):
        h2 = section.find("h2", class_="ttl")
        category = h2.get_text(strip=True) if h2 else ""
        for li in section.find_all("li"):
            name_div = li.find("div", class_="dl-link")
            pdf_a = li.find("a", class_="dl-link1")
            if name_div:
                reagents.append({
                    "reagent_name": name_div.get_text(strip=True),
                    "category": category,
                    "manufacturer": "カイノス",
                    "pdf_url": pdf_a.get("href", "") if pdf_a else "",
                    "source_url": url,
                })
    return reagents


# メーカースクレイパー登録（URL → スクレイパー関数）
MANUFACTURER_SCRAPERS = {
    "kainos_biochem": {
        "name": "カイノス（生化学）",
        "url": "https://www.kainos.co.jp/products/biochem/",
        "scraper": scrape_kainos_biochem,
    },
}


# ---------------------------------------------------------------------------
# 試薬 DB 構築
# ---------------------------------------------------------------------------

def build_reagent_db(
    output_dir: Path | None = None,
    manufacturers: list[str] | None = None,
    pmda_urls: list[str] | None = None,
) -> dict:
    """試薬DBを構築

    Args:
        output_dir: 出力ディレクトリ
        manufacturers: 取得するメーカーキーのリスト（Noneで全メーカー）
        pmda_urls: 手動指定のPMDA添付文書URL
    """
    output_dir = output_dir or Path("data")
    output_dir.mkdir(parents=True, exist_ok=True)
    reagent_dir = output_dir / "reagents"
    reagent_dir.mkdir(parents=True, exist_ok=True)

    session = requests.Session()
    now = datetime.now(timezone.utc)

    all_reagents: list[dict] = []

    # メーカーサイトから試薬一覧取得
    targets = manufacturers or list(MANUFACTURER_SCRAPERS.keys())
    for key in targets:
        if key not in MANUFACTURER_SCRAPERS:
            logger.warning("不明なメーカーキー: %s", key)
            continue
        info = MANUFACTURER_SCRAPERS[key]
        logger.info("試薬一覧取得: %s", info["name"])
        reagents = info["scraper"](session)
        all_reagents.extend(reagents)
        logger.info("  → %d件", len(reagents))
        time.sleep(REQUEST_INTERVAL)

    # PMDA URL から添付文書情報を取得（手動指定分）
    pmda_data: list[dict] = []
    if pmda_urls:
        for url in pmda_urls:
            logger.info("PMDA 添付文書取得: %s", url)
            try:
                detail = fetch_pmda_detail(url, session)
                detail["pmda_url"] = url
                pmda_data.append(detail)
            except Exception as e:
                logger.error("PMDA 取得エラー: %s - %s", url, e)
            time.sleep(REQUEST_INTERVAL)

    result = {
        "metadata": {
            "built_at": now.isoformat(),
            "total_reagents": len(all_reagents),
            "total_pmda": len(pmda_data),
            "manufacturers": targets,
        },
        "reagents": all_reagents,
        "pmda_details": pmda_data,
    }

    filepath = reagent_dir / "reagent_db.json"
    filepath.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    logger.info("試薬DB保存: %s (%d試薬, %dPMDA)", filepath, len(all_reagents), len(pmda_data))

    return result


def add_pmda_to_db(
    pmda_url: str,
    output_dir: Path | None = None,
) -> dict:
    """既存の試薬DBにPMDA添付文書情報を追加"""
    output_dir = output_dir or Path("data")
    db_path = output_dir / "reagents" / "reagent_db.json"

    if db_path.exists():
        db = json.loads(db_path.read_text(encoding="utf-8"))
    else:
        db = {"metadata": {}, "reagents": [], "pmda_details": []}

    session = requests.Session()
    logger.info("PMDA 添付文書取得: %s", pmda_url)
    detail = fetch_pmda_detail(pmda_url, session)
    detail["pmda_url"] = pmda_url

    # 重複チェック
    existing_urls = {d.get("pmda_url") for d in db.get("pmda_details", [])}
    if pmda_url in existing_urls:
        logger.info("既に登録済み: %s", pmda_url)
        # 上書き
        db["pmda_details"] = [d for d in db["pmda_details"] if d.get("pmda_url") != pmda_url]

    db["pmda_details"].append(detail)
    db["metadata"]["total_pmda"] = len(db["pmda_details"])
    db["metadata"]["last_updated"] = datetime.now(timezone.utc).isoformat()

    db_path.write_text(json.dumps(db, ensure_ascii=False, indent=2), encoding="utf-8")
    logger.info("PMDA追加完了: %s → %s", detail.get("product_name", ""), db_path)

    return detail
