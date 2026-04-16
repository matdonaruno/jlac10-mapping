"""外注検査 JLAC10 データ取得ツール CLI"""

import argparse
import json
import logging
import sys
from pathlib import Path

import requests

from .categories import CATEGORIES
from .scraper import check_update_needed, diff_report, scrape_all as srl_scrape_all
from .bml import check_bml_update_needed, scrape_all as bml_scrape_all
from .lsi import check_lsi_update_needed, scrape_all as lsi_scrape_all
from .merge import merge_all
from .jslm import scrape_all as jslm_scrape_all, check_jslm_update_needed
from .search import build_index, format_results
from .reagent import build_reagent_db, add_pmda_to_db
from .sop_parser import parse_sop, parse_sop_directory
from .converter import convert_tabular


def setup_logging(verbose: bool = False) -> None:
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )


# ---------------------------------------------------------------------------
# SRL
# ---------------------------------------------------------------------------

def cmd_srl(args: argparse.Namespace) -> int:
    category_ids = None
    if args.categories:
        category_ids = [int(c) for c in args.categories.split(",")]

    result = srl_scrape_all(
        category_ids=category_ids,
        output_dir=Path(args.output),
        use_cache=not args.no_cache,
        cache_max_age_hours=args.cache_age,
        check_update=args.check_update,
        force=args.force,
    )

    if result is None:
        print("\nSRL: 更新なし。スキップしました。")
        return 0

    meta = result["metadata"]
    print(f"\nSRL 取得完了: {meta['total_items']}件 ({meta['total_categories']}カテゴリ)")
    print(f"  サーバー: {meta['fetched_from_server']}件 / キャッシュ: {meta['served_from_cache']}件")
    if meta.get("last_update_date"):
        print(f"  Last Up Date: {meta['last_update_date']}")
    if meta["errors"]:
        print(f"  エラー: {len(meta['errors'])}件")
        return 1
    return 0


# ---------------------------------------------------------------------------
# BML
# ---------------------------------------------------------------------------

def cmd_bml(args: argparse.Namespace) -> int:
    result = bml_scrape_all(
        output_dir=Path(args.output),
        use_cache=not args.no_cache,
        cache_max_age_hours=args.cache_age,
        check_update=args.check_update,
    )

    if result is None:
        print("\nBML: 更新なし。スキップしました。")
        return 0

    meta = result["metadata"]
    print(f"\nBML 取得完了: {meta['total_items']}件 ({meta['total_categories']}カテゴリ)")
    print(f"  サーバー: {meta['fetched_from_server']}件 / キャッシュ: {meta['served_from_cache']}件")
    if meta["errors"]:
        print(f"  エラー: {len(meta['errors'])}件")
        return 1
    return 0


# ---------------------------------------------------------------------------
# LSI
# ---------------------------------------------------------------------------

def cmd_lsi(args: argparse.Namespace) -> int:
    result = lsi_scrape_all(
        output_dir=Path(args.output),
        use_cache=not args.no_cache,
        check_update=args.check_update,
        cache_max_age_hours=args.cache_age,
    )

    if result is None:
        print("\nLSI: 更新なし。スキップしました。")
        return 0

    meta = result["metadata"]
    print(f"\nLSI 取得完了: {meta['total_items']}件 ({meta['total_list_pages']}リストページ)")
    print(f"  詳細ページ サーバー: {meta['detail_pages_fetched']}件 / キャッシュ: {meta['detail_pages_cached']}件")
    if meta["errors"]:
        print(f"  エラー: {len(meta['errors'])}件")
        return 1
    return 0


# ---------------------------------------------------------------------------
# JSLM
# ---------------------------------------------------------------------------

def cmd_jslm(args: argparse.Namespace) -> int:
    result = jslm_scrape_all(
        output_dir=Path(args.output),
        check_update=args.check_update,
    )

    if result is None:
        print("\nJSLM: 更新なし。スキップしました。")
        return 0

    meta = result["metadata"]
    print(f"\nJSLM JLAC10マスター取得完了 (version: {meta['version']})")
    for k, v in meta["counts"].items():
        print(f"  {k}: {v}件")
    return 0


# ---------------------------------------------------------------------------
# 試薬DB
# ---------------------------------------------------------------------------

def cmd_reagent(args: argparse.Namespace) -> int:
    output_dir = Path(args.output)

    if args.pmda:
        # PMDA URLを追加
        detail = add_pmda_to_db(args.pmda, output_dir)
        print(f"\nPMDA添付文書追加:")
        print(f"  販売名: {detail.get('product_name', '')}")
        print(f"  使用目的: {detail.get('purpose', '')}")
        print(f"  測定原理: {detail.get('principle', '')[:100]}...")
        return 0

    # メーカー製品一覧取得
    result = build_reagent_db(output_dir=output_dir)
    meta = result["metadata"]
    print(f"\n試薬DB構築完了: {meta['total_reagents']}試薬, {meta['total_pmda']}PMDA添付文書")
    return 0


# ---------------------------------------------------------------------------
# SOP パーサー
# ---------------------------------------------------------------------------

def cmd_sop(args: argparse.Namespace) -> int:
    target = Path(args.path)
    output_dir = Path(args.output)

    if target.is_dir():
        result = parse_sop_directory(target, output_dir)
        meta = result["metadata"]
        print(f"\nSOP パース完了: {meta['total_files']}件")
        print(f"  測定法あり: {meta['with_method']}件")
        print(f"  エラー: {meta['errors']}件")
        print(f"  出力: {output_dir}/sop_parsed.json")
    elif target.is_file():
        info = parse_sop(target)
        print(f"\nSOP パース結果: {target.name}")
        print(f"  検査項目: {info['test_item'][:60] or '(未検出)'}")
        print(f"  測定法:   {info['method_summary'][:80] or '(未検出)'}")
        print(f"  試薬:     {info['reagent'][:80] or '(未検出)'}")
        print(f"  装置:     {info['instrument'][:60] or '(未検出)'}")
        print(f"  検体:     {info['specimen'][:60] or '(未検出)'}")
        if info.get("error"):
            print(f"  エラー:   {info['error']}")

        # JSON出力
        out = output_dir / f"sop_{target.stem}.json"
        output_dir.mkdir(parents=True, exist_ok=True)
        out.write_text(json.dumps(info, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"  出力: {out}")
    else:
        print(f"ファイル/ディレクトリが見つかりません: {target}", file=sys.stderr)
        return 1
    return 0


# ---------------------------------------------------------------------------
# Convert (Excel/CSV → JSON)
# ---------------------------------------------------------------------------

def cmd_convert(args: argparse.Namespace) -> int:
    filepath = Path(args.input)

    # column_map を構築
    column_map: dict[str, str] = {}
    if args.col_item:
        column_map["item_name"] = args.col_item
    if args.col_jlac10:
        column_map["jlac10"] = args.col_jlac10
    if args.col_abbr:
        column_map["abbreviation"] = args.col_abbr
    if args.col_std_name:
        column_map["jlac10_standard_name"] = args.col_std_name

    if "item_name" not in column_map or "jlac10" not in column_map:
        print(
            "エラー: --col-item と --col-jlac10 は必須です",
            file=sys.stderr,
        )
        return 1

    output_path = Path(args.output_file) if args.output_file else None

    try:
        result = convert_tabular(
            filepath=filepath,
            column_map=column_map,
            hospital=args.hospital,
            sheet_name=args.sheet,
            skip_rows=args.skip_rows,
            output_path=output_path,
        )
    except (FileNotFoundError, ValueError) as e:
        print(f"エラー: {e}", file=sys.stderr)
        return 1

    meta = result["metadata"]
    out = output_path or filepath.with_suffix(".json")
    print(f"\n変換完了: {meta['total_items']}件")
    print(f"  病院: {meta['hospital'] or '(未指定)'}")
    print(f"  ソース: {meta['source_file']}")
    print(f"  出力: {out}")
    return 0


# ---------------------------------------------------------------------------
# 統合・差分・チェック・一覧
# ---------------------------------------------------------------------------

def cmd_merge(args: argparse.Namespace) -> int:
    result = merge_all(output_dir=Path(args.output))
    meta = result["metadata"]
    print(f"\n統合完了: {meta['total_unique_jlac10']}件のユニーク JLAC10")
    print(f"  データソース: {meta['sources_available']}")
    counts = meta["by_source_count"]
    print(f"  3社共通: {counts['all_three']}件 / 2社: {counts['two_sources']}件")
    print(f"  SRLのみ: {counts['srl_only']}件 / BMLのみ: {counts['bml_only']}件 / LSIのみ: {counts['lsi_only']}件")
    return 0


def cmd_check(args: argparse.Namespace) -> int:
    output_dir = Path(args.output)
    output_dir.mkdir(parents=True, exist_ok=True)
    session = requests.Session()

    import time
    from .scraper import REQUEST_INTERVAL

    # SRL
    srl_needed, srl_remote, srl_local = check_update_needed(output_dir, session)
    print(f"[SRL] Last Up Date: {srl_remote} (前回: {srl_local or '未取得'})")
    print(f"  → {'更新あり' if srl_needed else '更新なし'}")
    time.sleep(REQUEST_INTERVAL)

    # BML
    bml_needed, bml_remote, bml_local = check_bml_update_needed(output_dir, session)
    print(f"[BML] 最新掲載日:  {bml_remote} (前回: {bml_local or '未取得'})")
    print(f"  → {'更新あり' if bml_needed else '更新なし'}")
    time.sleep(REQUEST_INTERVAL)

    # LSI
    lsi_needed, lsi_remote, lsi_local = check_lsi_update_needed(output_dir, session)
    print(f"[LSI] 掲載日:      {lsi_remote} (前回: {lsi_local or '未取得'})")
    print(f"  → {'更新あり' if lsi_needed else '更新なし'}")
    time.sleep(REQUEST_INTERVAL)

    # JSLM
    jslm_needed, jslm_remote, jslm_local = check_jslm_update_needed(output_dir, session)
    print(f"[JSLM] コード表:   version {jslm_remote} (前回: {jslm_local or '未取得'})")
    print(f"  → {'更新あり' if jslm_needed else '更新なし'}")

    return 0


def cmd_diff(args: argparse.Namespace) -> int:
    old_path = Path(args.old)
    new_path = Path(args.new)

    if not old_path.exists():
        print(f"ファイルが見つかりません: {old_path}", file=sys.stderr)
        return 1
    if not new_path.exists():
        print(f"ファイルが見つかりません: {new_path}", file=sys.stderr)
        return 1

    report = diff_report(old_path, new_path)
    s = report["summary"]
    print(f"\n差分レポート:")
    print(f"  追加: {s['added']}件 / 削除: {s['removed']}件 / 変更: {s['changed']}件 / 変更なし: {s['unchanged']}件")

    if args.output_report:
        out = Path(args.output_report)
        out.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"  レポート保存: {out}")
    return 0


def cmd_search(args: argparse.Namespace) -> int:
    data_dir = Path(args.output)
    index = build_index(data_dir)
    print(f"インデックス構築完了: {len(index.entries)}件")

    if args.query:
        # 引数で検索
        results = index.search(args.query, max_results=args.max)
        print(f"\n検索: 「{args.query}」 → {len(results)}件ヒット")
        print(format_results(results))
    else:
        # 対話モード
        print("検索クエリを入力してください（quit で終了）:\n")
        while True:
            try:
                query = input("検索> ").strip()
            except (EOFError, KeyboardInterrupt):
                print()
                break
            if not query or query.lower() in ("quit", "exit", "q"):
                break
            results = index.search(query, max_results=args.max)
            print(f"\n「{query}」 → {len(results)}件ヒット")
            print(format_results(results))
            print()
    return 0


def cmd_list(args: argparse.Namespace) -> int:
    current_group = ""
    for cat in CATEGORIES:
        if cat.group != current_group:
            current_group = cat.group
            print(f"\n[{current_group}]")
        print(f"  {cat.id:>3d}: {cat.name}")
    return 0


# ---------------------------------------------------------------------------
# メイン
# ---------------------------------------------------------------------------

def _add_cache_args(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--no-cache", action="store_true", help="キャッシュ無視")
    parser.add_argument("--cache-age", type=float, default=24.0, help="キャッシュ有効期限(時間)")
    parser.add_argument("-o", "--output", default="data", help="出力ディレクトリ")


def main() -> None:
    parser = argparse.ArgumentParser(
        prog="srl-scraper",
        description="外注検査 JLAC10 データ取得ツール（SRL / BML / LSI メディエンス）",
    )
    parser.add_argument("-v", "--verbose", action="store_true", help="詳細ログ出力")
    sub = parser.add_subparsers(dest="command")

    # srl
    p_srl = sub.add_parser("srl", help="SRL検査案内からデータ取得")
    _add_cache_args(p_srl)
    p_srl.add_argument("-c", "--categories", help="カテゴリIDをカンマ区切り (例: 1,2,3)")
    p_srl.add_argument("--check-update", action="store_true", help="Last Up Date 確認、更新時のみ取得")
    p_srl.add_argument("--force", action="store_true", help="check-update時でも強制取得")

    # bml
    p_bml = sub.add_parser("bml", help="BML検査案内からデータ取得")
    _add_cache_args(p_bml)
    p_bml.add_argument("--check-update", action="store_true", help="new/changeの掲載日を確認、更新時のみ取得")

    # lsi
    p_lsi = sub.add_parser("lsi", help="LSIメディエンス検査案内からデータ取得")
    _add_cache_args(p_lsi)
    p_lsi.add_argument("--check-update", action="store_true", help="掲載日を確認、更新時のみ取得")

    # jslm
    p_jslm = sub.add_parser("jslm", help="JSLM JLAC10コード表を取得")
    p_jslm.add_argument("-o", "--output", default="data", help="出力ディレクトリ")
    p_jslm.add_argument("--check-update", action="store_true", help="版数確認、更新時のみ取得")

    # reagent
    p_reagent = sub.add_parser("reagent", help="試薬DB構築（メーカー製品 + PMDA添付文書）")
    p_reagent.add_argument("-o", "--output", default="data", help="出力ディレクトリ")
    p_reagent.add_argument("--pmda", help="PMDA添付文書URLを追加")

    # sop
    p_sop = sub.add_parser("sop", help="SOP(Word/PDF)から測定法・試薬情報を抽出")
    p_sop.add_argument("path", help="SOPファイル(.docx/.pdf) または ディレクトリ")
    p_sop.add_argument("-o", "--output", default="data", help="出力ディレクトリ")

    # search
    p_search = sub.add_parser("search", help="検査項目をあいまい検索")
    p_search.add_argument("query", nargs="?", default=None, help="検索文字列（省略で対話モード）")
    p_search.add_argument("-o", "--output", default="data", help="データディレクトリ")
    p_search.add_argument("-n", "--max", type=int, default=10, help="最大結果数 (default: 10)")

    # merge
    p_merge = sub.add_parser("merge", help="3社のデータをJLAC10で統合")
    p_merge.add_argument("-o", "--output", default="data", help="出力ディレクトリ")

    # check
    p_check = sub.add_parser("check", help="全ソースの更新状況を確認")
    p_check.add_argument("-o", "--output", default="data", help="出力ディレクトリ")

    # diff
    p_diff = sub.add_parser("diff", help="2つのJSONファイルの差分表示")
    p_diff.add_argument("old", help="古いJSON")
    p_diff.add_argument("new", help="新しいJSON")
    p_diff.add_argument("-o", "--output-report", help="差分レポート出力先")

    # convert
    p_convert = sub.add_parser("convert", help="院内検査マスタ(Excel/CSV)をJSONに変換")
    p_convert.add_argument("input", help="入力ファイル (.xlsx / .csv)")
    p_convert.add_argument("-o", "--output-file", default=None, help="出力JSONパス (省略で {入力ファイル名}.json)")
    p_convert.add_argument("--hospital", default="", help="病院名")
    p_convert.add_argument("--col-item", required=True, help="検査項目名の列 (A, 1, またはヘッダ名)")
    p_convert.add_argument("--col-jlac10", required=True, help="JLAC10の列 (A, 1, またはヘッダ名)")
    p_convert.add_argument("--col-abbr", default=None, help="略称の列 (A, 1, またはヘッダ名)")
    p_convert.add_argument("--col-std-name", default=None, help="JLAC10標準名称の列 (A, 1, またはヘッダ名)")
    p_convert.add_argument("--sheet", default=None, help="Excelシート名 (省略で最初のシート)")
    p_convert.add_argument("--skip-rows", type=int, default=1, help="スキップするヘッダ行数 (default: 1)")

    # list
    sub.add_parser("list", help="SRLカテゴリ一覧")

    args = parser.parse_args()
    setup_logging(args.verbose)

    if not args.command:
        parser.print_help()
        sys.exit(1)

    commands = {
        "srl": cmd_srl,
        "bml": cmd_bml,
        "lsi": cmd_lsi,
        "jslm": cmd_jslm,
        "reagent": cmd_reagent,
        "sop": cmd_sop,
        "search": cmd_search,
        "merge": cmd_merge,
        "check": cmd_check,
        "diff": cmd_diff,
        "convert": cmd_convert,
        "list": cmd_list,
    }
    sys.exit(commands[args.command](args))
