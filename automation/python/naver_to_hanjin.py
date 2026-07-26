#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
naver_to_hanjin.py
네이버 스마트스토어 주문 엑셀/CSV → 한진 원클릭 택배 업로드 양식 변환기.

사용법:
    python naver_to_hanjin.py 주문.xlsx
    python naver_to_hanjin.py 주문.csv -o 한진송장.csv
    python naver_to_hanjin.py 주문.xlsx --config config.yaml --xlsx

설정(컬럼 매핑)은 config.yaml 에서 조정합니다.
의존성: pandas, pyyaml, openpyxl  (requirements.txt 참고)
"""
import argparse
import re
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
import yaml


def load_config(path: Path) -> dict:
    with open(path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def read_orders(path: Path) -> pd.DataFrame:
    """엑셀/CSV 주문 파일을 읽어 DataFrame으로. 헤더 공백은 정리."""
    if path.suffix.lower() in (".xlsx", ".xls"):
        df = pd.read_excel(path, dtype=str)
    else:
        # 한글 CSV 인코딩 자동 대응
        for enc in ("utf-8-sig", "cp949", "euc-kr", "utf-8"):
            try:
                df = pd.read_csv(path, dtype=str, encoding=enc)
                break
            except (UnicodeDecodeError, UnicodeError):
                continue
        else:
            raise SystemExit(f"CSV 인코딩을 읽지 못했습니다: {path}")
    df.columns = [str(c).strip() for c in df.columns]
    return df.fillna("")


def find_col(df: pd.DataFrame, name: str):
    """헤더를 정확일치 우선, 없으면 부분 포함으로 찾아 실제 컬럼명 반환."""
    if not name:
        return None
    if name in df.columns:
        return name
    for c in df.columns:
        if name in c or c in name:
            return c
    return None


def get(row: pd.Series, df: pd.DataFrame, naver_cols: dict, key: str) -> str:
    col = find_col(df, naver_cols.get(key, ""))
    if col is None:
        return ""
    val = row.get(col, "")
    return "" if pd.isna(val) else str(val).strip()


def normalize_phone(v: str) -> str:
    """숫자만 뽑아 010-xxxx-xxxx / 지역번호 형태로 정규화."""
    d = re.sub(r"[^0-9]", "", str(v or ""))
    if not d:
        return ""
    if len(d) == 10 and d[0] == "1":       # 1012345678 → 01012345678 (앞 0 소실 보정)
        d = "0" + d
    if len(d) == 11:
        return f"{d[0:3]}-{d[3:7]}-{d[7:11]}"
    if len(d) == 10:
        if d.startswith("02"):
            return f"{d[0:2]}-{d[2:6]}-{d[6:10]}"
        return f"{d[0:3]}-{d[3:6]}-{d[6:10]}"
    return d


def item_label(name: str, opt: str, qty) -> str:
    try:
        q = int(float(qty))
    except (ValueError, TypeError):
        q = 1
    q = q if q > 0 else 1
    s = (name or "").strip()
    if opt:
        s += f" ({opt.strip()})"
    if q > 1:
        s += f" x{q}"
    return s


def convert(df: pd.DataFrame, cfg: dict) -> pd.DataFrame:
    nc = cfg["naver_cols"]
    hanjin_cols = cfg["hanjin_cols"]
    mapping = cfg["mapping"]
    combine = cfg.get("combine_by_receiver", True)

    # 1) 네이버 행 → 정규화된 dict 목록
    records = []
    for _, row in df.iterrows():
        receiver = get(row, df, nc, "receiver")
        address = get(row, df, nc, "address")
        if not receiver and not address:
            continue  # 빈 줄
        records.append({
            "receiver": receiver,
            "address": address,
            "phone": get(row, df, nc, "receiver_phone"),
            "phone2": get(row, df, nc, "receiver_phone2"),
            "zipcode": get(row, df, nc, "zipcode"),
            "product_name": get(row, df, nc, "product_name"),
            "option": get(row, df, nc, "option"),
            "quantity": get(row, df, nc, "quantity"),
            "delivery_memo": get(row, df, nc, "delivery_memo"),
        })

    # 2) 수취인 기준 병합
    if combine:
        merged = {}
        order = []
        for r in records:
            key = (r["receiver"].strip(),
                   re.sub(r"\s+", "", r["address"]),
                   normalize_phone(r["phone"]))
            if key not in merged:
                m = dict(r)
                m["_items"] = []
                m["_qty"] = 0
                merged[key] = m
                order.append(key)
            qty = int(float(r["quantity"])) if str(r["quantity"]).strip() else 1
            merged[key]["_items"].append(item_label(r["product_name"], r["option"], qty))
            merged[key]["_qty"] += qty
        records = []
        for key in order:
            m = merged[key]
            m["_item_name"] = " + ".join(m["_items"])
            m["quantity"] = m["_qty"]
            records.append(m)

    # 3) 매핑 규칙대로 한진 행 생성
    def compute(fn: str, r: dict) -> str:
        if fn == "phone":
            return normalize_phone(r["phone"])
        if fn == "phone2":
            return normalize_phone(r["phone2"])
        if fn == "item_name":
            if "_item_name" in r:
                return r["_item_name"][:100]
            return item_label(r["product_name"], r["option"], r["quantity"])[:100]
        return ""

    out_rows = []
    for r in records:
        row_out = {}
        for col in hanjin_cols:
            rule = mapping.get(col, {})
            if "const" in rule:
                row_out[col] = rule["const"]
            elif "from" in rule:
                row_out[col] = r.get(rule["from"], "")
            elif "fn" in rule:
                row_out[col] = compute(rule["fn"], r)
            else:
                row_out[col] = ""
        out_rows.append(row_out)

    return pd.DataFrame(out_rows, columns=hanjin_cols)


def main():
    ap = argparse.ArgumentParser(description="네이버 주문 → 한진 원클릭 양식 변환")
    ap.add_argument("input", help="네이버 주문 파일 (.xlsx / .csv)")
    ap.add_argument("-o", "--output", help="출력 파일 경로 (기본: 한진송장_날짜.csv)")
    ap.add_argument("--config", default=str(Path(__file__).parent / "config.yaml"),
                    help="매핑 설정 yaml (기본: 옆의 config.yaml)")
    ap.add_argument("--xlsx", action="store_true", help="CSV 대신 XLSX로 출력")
    args = ap.parse_args()

    in_path = Path(args.input)
    if not in_path.exists():
        raise SystemExit(f"입력 파일 없음: {in_path}")

    cfg = load_config(Path(args.config))
    df = read_orders(in_path)
    result = convert(df, cfg)

    if args.output:
        out_path = Path(args.output)
    else:
        stamp = datetime.now().strftime("%Y%m%d_%H%M")
        ext = "xlsx" if args.xlsx else "csv"
        out_path = in_path.parent / f"한진송장_{stamp}.{ext}"

    if args.xlsx or out_path.suffix.lower() == ".xlsx":
        result.to_excel(out_path, index=False)
    else:
        # 한진/엑셀 한글 호환: UTF-8 BOM
        result.to_csv(out_path, index=False, encoding="utf-8-sig")

    print(f"✅ 변환 완료: 송장 {len(result)}건 → {out_path}")


if __name__ == "__main__":
    main()
