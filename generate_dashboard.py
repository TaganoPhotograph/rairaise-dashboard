#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
らいらいず 個人成績・段位進捗ダッシュボード生成スクリプト（複数ペア対応版）

使い方:
    python3 generate_dashboard.py <入力xlsxパス> <出力htmlパス>

新xlsx仕様（複数ペア対応版）:
- 「一覧」シート: 5行目がヘッダー、6行目以降がデータ。列は
  名前キー(A) / 表示名(B) / はなまるNO(C) / 審査段位(D) / 平均着順(E) /
  ラス回避率(F) / トップ回数(G) / 実践回数(H) / ポイント(I)
- 各メンバーの個人タブ（シート名 = 一覧のA列キーと一致）
  - 各「審査ペア」は連続する2行で構成される（上=素点、下=着順）
  - 1組目: 4-5行目
  - 2組目: 6-7行目
  - 3組目: 8-9行目
  - ...（以降 2行ずつ増えていく。人によりペア数が異なる）
  - 各ペアの F列（列6）に段位ラベルがあることが多い（無い場合もある）
  - 各ペアの J〜AM列（列10〜39）が 実績1〜30（該当ペアの各半荘の得点/着順）
  - BY5:CB18（列77〜80、行5〜18）に段位別の昇段基準テーブル
- 「原本」シートは無視する。

計算ルール（★ユーザー指定）:
- 累計成績 = 該当タブ内の【全ペア】の全対局データを合算した集計
- 期間内成績 = 該当タブ内の【一番下（最下段）のデータ入りペア】の集計
- 段位判定:
  - 最下段ペアの段位ラベル = 一覧の審査段位 → そのペアが現在の審査
  - 段位ラベルが空欄 → 一覧の審査段位を採用（同じ段位の継続とみなす）
  - ラベルと一覧が食い違う → その段位を突破して新段位に上がり、新段位ではまだ0半荘
    （その旨を確認事項として報告する）
"""

import sys, json, math, datetime
import pandas as pd
import openpyxl

RANK_ORDER = ["5級","4級","3級","2級","1級","初段","二段","三段","四段","五段","六段","七段","八段","九段"]


def to_num(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        if isinstance(v, float) and math.isnan(v):
            return None
        return v
    return None


def extract_pairs(ws):
    """個人シートから全ペアを抽出。データが1件でもあるペアのみ返す。"""
    pairs = []
    consecutive_empty = 0
    for pair_idx in range(60):  # 十分な数の枠を確認
        r_score = 4 + pair_idx * 2
        r_place = r_score + 1
        rank_label = ws.cell(row=r_score, column=6).value
        rank_label = str(rank_label).strip() if rank_label else None
        games = []
        for c in range(10, 40):
            s = to_num(ws.cell(row=r_score, column=c).value)
            p = to_num(ws.cell(row=r_place, column=c).value)
            if s is not None and p in (1, 2, 3, 4):
                games.append({"score": s, "place": int(p)})
        if not games:
            if not rank_label:
                consecutive_empty += 1
                if consecutive_empty >= 3:
                    break  # 3枠連続で完全に空なら終端とみなす
            else:
                consecutive_empty = 0
            continue
        consecutive_empty = 0
        pairs.append({"rank": rank_label, "games": games})
    return pairs


def compute_window_stats(games):
    played = len(games)
    if played == 0:
        return {"played": 0, "avg_place": None, "avg_score": 0.0, "place_sum": 0,
                "last_avoid_rate": None, "total_score": 0}
    place_sum = sum(g["place"] for g in games)
    total_score = sum(g["score"] for g in games)
    avoid = sum(1 for g in games if g["place"] != 4)
    return {
        "played": played,
        "avg_place": round(place_sum / played, 3),
        "avg_score": round(total_score / played, 2),
        "place_sum": place_sum,
        "last_avoid_rate": round(avoid / played, 3),
        "total_score": total_score,
    }


def extract(xlsx_path):
    df = pd.read_excel(xlsx_path, sheet_name="一覧", header=None)
    # 一覧シートの想定列数(A~I の9列)に限定して読む。
    # エクスポート経路によっては末尾に書式のみの空列が付与されることがあるため、
    # 9列を超える分は無視する（データとしては使われていない）。
    raw_rows = df.iloc[5:, 0:9].values.tolist()
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)

    # 段位基準テーブル（どのシートでも共通のはずなので、"一覧"・"原本"以外の最初のシートから読む）
    ref_sheet_name = next(s for s in wb.sheetnames if s not in ("一覧", "原本"))
    ref_ws = wb[ref_sheet_name]
    rank_req = {}
    for i, rank in enumerate(RANK_ORDER):
        r = 5 + i
        rank_req[rank] = {
            "games": ref_ws.cell(row=r, column=78).value,
            "place": ref_ws.cell(row=r, column=79).value,
            "avg": ref_ws.cell(row=r, column=80).value,
        }

    players = []
    seen = set()
    warnings = []
    indirect_fallback_count = 0

    for r in raw_rows:
        key, disp_name, hanamaru_no, rank, avg, lastavoid, top, games, pts = r
        if pd.isna(key) or str(key).strip() == '':
            continue
        key = str(key).strip()

        # ---- INDIRECT数式キャッシュ切れフォールバック ----
        # 「一覧」シートの表示名(B列)・審査段位(D列)はINDIRECT数式のため、
        # ODS経由の変換などでキャッシュが失われると空欄(NaN)で返ってくることがある。
        # その場合は個人タブ自身の C4(表示名)・C12(現状段位) から直接読み取って補完する。
        # （INDIRECTが本来参照している先そのものを読むだけで、判定ロジック・基準値は変更しない）
        used_fallback = False
        if key in wb.sheetnames:
            tab_ws = wb[key]
            if pd.isna(disp_name) or str(disp_name).strip() == '':
                c4 = tab_ws.cell(row=4, column=3).value
                if c4 not in (None, ''):
                    disp_name = c4
                    used_fallback = True
            if pd.isna(rank) or str(rank).strip() == '':
                c12 = tab_ws.cell(row=12, column=3).value
                if c12 not in (None, ''):
                    rank = c12
                    used_fallback = True
        if used_fallback:
            indirect_fallback_count += 1

        disp_name = None if pd.isna(disp_name) else str(disp_name).strip()
        name = disp_name if (disp_name and disp_name != "0") else key

        rank = str(rank).strip() if not pd.isna(rank) else ""
        top_n = int(top) if not pd.isna(top) else 0

        dedup_key = (key, rank)
        if dedup_key in seen:
            continue
        seen.add(dedup_key)

        exam_target = rank if rank in RANK_ORDER else None

        if key not in wb.sheetnames:
            warnings.append(f"一覧に「{key}」がありますが、同名の個人タブが見つかりません。")
            pairs = []
        else:
            pairs = extract_pairs(wb[key])

        # ---- 累計成績: 全ペアを合算 ----
        all_games = [g for p in pairs for g in p["games"]]
        career = compute_window_stats(all_games)
        career_top = sum(1 for g in all_games if g["place"] == 1)
        career["top_count"] = career_top

        # ---- 期間内成績: 最下段ペア ----
        exam = None
        if exam_target:
            if not pairs:
                exam = _make_exam(exam_target, rank_req[exam_target], [])
            else:
                last = pairs[-1]
                last_rank = last["rank"]
                if last_rank is None or last_rank == exam_target:
                    # 継続 or ラベル省略 → 最下段ペアをそのまま期間内として採用
                    used_rank = exam_target
                    exam = _make_exam(used_rank, rank_req[used_rank], last["games"])
                else:
                    # 最下段ペアの段位ラベルが一覧と食い違う → 新段位に上がったばかりで0半荘
                    warnings.append(
                        f"「{name}」: 最下段ペアのラベル({last_rank})と一覧の審査段位({exam_target})が"
                        f"食い違うため、新段位0半荘スタートとして計算しました。"
                    )
                    exam = _make_exam(exam_target, rank_req[exam_target], [])
        else:
            exam = {"rank": None, "status": "unranked"}

        # 着順別の平均素点（シミュレーターの収支推定に使う）
        place_avg_score = {}
        for k in (1, 2, 3, 4):
            ss = [g["score"] for g in all_games if g["place"] == k]
            place_avg_score[k] = {"avg": round(sum(ss) / len(ss), 1) if ss else None, "n": len(ss)}

        players.append({
            "key": key,
            "name": name,
            "rank": rank,
            "rank_value": RANK_ORDER.index(rank) if rank in RANK_ORDER else -1,
            "career": {
                "games": career["played"],
                "avg_place": career["avg_place"],
                "last_avoid_rate": career["last_avoid_rate"],
                "points": round(career["total_score"], 1),
                "top_count": career_top,
                # 全半荘の推移（古い順）。成績推移グラフに使う
                "places": [g["place"] for g in all_games],
                "place_avg_score": place_avg_score,
            },
            "exam": exam,
        })

    if indirect_fallback_count > 0:
        warnings.insert(0,
            f"[自動補完・要確認不要] {indirect_fallback_count}名分の表示名/審査段位をINDIRECT数式"
            f"キャッシュ切れのため個人タブ(C4/C12)から直接補完しました。"
        )

    # サークル全体の着順別平均素点（本人のデータが少ないときの補完用）
    circle_place_avg = {}
    for k in (1, 2, 3, 4):
        tot = sum((p["career"]["place_avg_score"][k]["avg"] or 0) * p["career"]["place_avg_score"][k]["n"] for p in players)
        cnt = sum(p["career"]["place_avg_score"][k]["n"] for p in players)
        circle_place_avg[k] = round(tot / cnt, 1) if cnt else None

    return {"rank_order": RANK_ORDER, "rank_req": rank_req, "players": players,
            "circle_place_avg": circle_place_avg}, warnings


REACH_REMAINING = 3  # 「昇段リーチ」とみなす残り半荘数（基準クリア中かつ残りこの数以下）


def load_prev_data(html_path):
    """前回生成した index.html から埋め込みデータを読み出す。無ければ None。"""
    import os, re
    if not html_path or not os.path.exists(html_path):
        return None
    s = open(html_path, encoding="utf-8").read()
    m = re.search(r'<script id="rawData" type="application/json">(.*?)</script>', s, re.S)
    gen = re.search(r'<div class="updated">(.*?)</div>', s)
    if not m:
        return None
    d = json.loads(m.group(1))
    d["_generated_at"] = gen.group(1) if gen else None
    return d


def compute_changes(data, prev):
    """前回との差分（昇段者・新メンバー）と現在の注目（合格ライン到達・昇段リーチ）を作る。"""
    players = data["players"]
    ready = [p["key"] for p in players if p["exam"] and p["exam"].get("status") == "ready"]
    reach = [p["key"] for p in players if p["exam"] and p["exam"].get("status") == "on_track"
             and 0 < p["exam"].get("remaining_games", 0) <= REACH_REMAINING]
    changes = {"ready": ready, "reach": reach, "promoted": [], "new": [], "since": None}
    if not prev:
        return changes

    P = {p["key"]: p for p in prev.get("players", [])}
    games_changed = any(p["key"] not in P or p["career"]["games"] != P[p["key"]]["career"]["games"]
                        for p in players)
    if not games_changed and prev.get("changes"):
        # 同じデータでの再生成 → 前回の「前回比」をそのまま引き継ぐ
        old = prev["changes"]
        changes.update({"promoted": old.get("promoted", []), "new": old.get("new", []),
                        "since": old.get("since")})
        return changes

    changes["since"] = (prev.get("_generated_at") or "").replace(" 更新", "")
    for p in players:
        q = P.get(p["key"])
        if q is None:
            changes["new"].append(p["key"])
        elif p["rank_value"] > q.get("rank_value", -1) and q.get("rank_value", -1) >= 0:
            changes["promoted"].append({"key": p["key"], "from": q["rank"], "to": p["rank"]})
    return changes


def _make_exam(rank, req, games):
    stats = compute_window_stats(games)
    played = stats["played"]
    required_games = req["games"]
    required_place = req["place"]
    required_avg = req["avg"]
    remaining_games = max(required_games - played, 0)
    place_clear = (stats["avg_place"] is not None) and (stats["avg_place"] <= required_place)
    avg_clear = played > 0 and stats["avg_score"] >= required_avg
    total_required = required_avg * required_games
    total_remaining = max(total_required - stats["total_score"], 0)
    pace_needed = round(total_remaining / remaining_games, 1) if remaining_games > 0 else None

    if played == 0:
        status = "not_started"
    elif played >= required_games:
        status = "ready" if (place_clear and avg_clear) else "reset_pending"
    else:
        status = "on_track" if (place_clear and avg_clear) else "behind"

    return {
        "rank": rank,
        "rank_value": RANK_ORDER.index(rank),
        "required_games": required_games,
        "required_place": required_place,
        "required_avg": required_avg,
        "played_games": played,
        "avg_place": stats["avg_place"],
        "avg_score": stats["avg_score"],
        "place_sum": stats["place_sum"],
        "last_avoid_rate": stats["last_avoid_rate"],
        "remaining_games": remaining_games,
        "place_clear": place_clear,
        "avg_clear": avg_clear,
        "total_current": round(stats["total_score"], 1),
        "total_required": round(total_required, 1),
        "total_remaining": round(total_remaining, 1),
        "pace_needed": pace_needed,
        "status": status,
        # 期間内の各半荘（古い順）。進捗メーターの着順表示に使う
        "places": [g["place"] for g in games],
        "scores": [round(g["score"], 1) for g in games],
    }


TEMPLATE = r"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>らいらいず 個人成績・段位進捗</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Shippori+Mincho:wght@500;700;800&family=Zen+Kaku+Gothic+New:wght@400;500;700;900&display=swap" rel="stylesheet">
<style>
  :root{
    --felt-1:#0b3d2e; --felt-2:#0e2f26; --felt-edge:#062018;
    --tile:#f6efdd; --tile-edge:#cdbb8e;
    --ink:#20281f; --ink-soft:#5b6357;
    --red:#a8322c; --green-dragon:#1f6b3a; --gold:#b8860b;
    --amber:#b3781f;
    --shadow: 0 6px 16px rgba(0,0,0,.35);
  }
  *{box-sizing:border-box;}
  html,body{margin:0;padding:0;}
  body{
    font-family:'Zen Kaku Gothic New', sans-serif;
    color:var(--ink);
    background:
      radial-gradient(circle at 15% 10%, rgba(255,255,255,.03), transparent 40%),
      repeating-linear-gradient(45deg, rgba(255,255,255,0.015) 0 2px, transparent 2px 10px),
      linear-gradient(160deg, var(--felt-1), var(--felt-2) 60%, var(--felt-edge));
    min-height:100vh;
    padding-bottom:80px;
  }
  .wrap{max-width:1040px;margin:0 auto;padding:36px 20px 0;}

  header.plaque{text-align:center;padding:38px 20px 30px;}
  .kanban{
    display:inline-block; background:var(--tile); border:1px solid var(--tile-edge);
    border-radius:10px; box-shadow: var(--shadow), inset 0 0 0 3px rgba(255,255,255,.4);
    padding:22px 48px; position:relative;
  }
  .kanban::before, .kanban::after{content:"";position:absolute;top:8px;bottom:8px;width:1px;background:rgba(0,0,0,.08);}
  .kanban::before{left:10px;} .kanban::after{right:10px;}
  .kanban h1{font-family:'Shippori Mincho', serif;font-weight:800;font-size:2.6rem;letter-spacing:.14em;margin:0;color:var(--ink);}
  .kanban .sub{margin-top:6px;font-size:.82rem;letter-spacing:.2em;color:var(--ink-soft);}
  .updated{margin-top:14px;font-size:.78rem;color:rgba(255,255,255,.55);letter-spacing:.05em;}

  .statbar{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin:0 0 34px;}
  .stat{background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.12);border-radius:10px;padding:14px 10px;text-align:center;color:#eef3ee;}
  .stat .n{font-family:'Shippori Mincho', serif;font-weight:700;font-size:1.5rem;color:#f6efdd;}
  .stat .l{font-size:.66rem;letter-spacing:.1em;color:rgba(255,255,255,.55);margin-top:2px;}

  h2.section-title{
    font-family:'Shippori Mincho', serif; color:#f6efdd; font-size:1.05rem;
    letter-spacing:.2em; text-align:center; margin:0 0 18px; opacity:.85;
  }
  h2.section-title .asof{
    display:block; font-family:'Zen Kaku Gothic New',sans-serif; font-weight:400;
    letter-spacing:.05em; font-size:.68rem; color:rgba(255,255,255,.5);
    margin-top:6px;
  }

  .picker{display:flex; gap:10px; flex-wrap:wrap; align-items:center; justify-content:center; margin-bottom:20px;}
  .combo{position:relative; width:100%; max-width:420px;}
  .combo input{
    width:100%; background:var(--tile); border:2px solid var(--tile-edge); color:var(--ink);
    font-family:'Zen Kaku Gothic New',sans-serif; font-weight:700; font-size:16px;
    padding:12px 44px 12px 18px; border-radius:24px; outline:none; -webkit-appearance:none; appearance:none;
  }
  .combo input::placeholder{color:#8a8574; font-weight:500;}
  .combo input:focus{border-color:var(--gold); box-shadow:0 0 0 3px rgba(184,134,11,.25);}
  .combo::after{content:"⌕"; position:absolute; right:16px; top:50%; transform:translateY(-50%); font-size:1.2rem; color:var(--ink-soft); pointer-events:none;}
  .combo ul{
    position:absolute; left:0; right:0; top:calc(100% + 6px); z-index:20; margin:0; padding:6px; list-style:none;
    background:var(--tile); border:1px solid var(--tile-edge); border-radius:14px; box-shadow:var(--shadow);
    max-height:min(60vh, 360px); overflow-y:auto; -webkit-overflow-scrolling:touch;
  }
  .combo li{display:flex; align-items:center; justify-content:space-between; gap:10px; padding:10px 12px; border-radius:10px; cursor:pointer; font-weight:700; color:var(--ink);}
  .combo li[aria-selected="true"], .combo li:hover{background:#efe6cd;}
  .combo li .badge{font-size:.66rem; padding:3px 9px; border-radius:20px; color:#fff; font-weight:700; white-space:nowrap;}
  .combo li.none{color:var(--ink-soft); font-weight:500; cursor:default; background:none;}
  .picker.first .combo input{border-color:var(--gold); animation:pulse 2s ease-in-out infinite;}
  @keyframes pulse{0%,100%{box-shadow:0 0 0 0 rgba(184,134,11,.0);}50%{box-shadow:0 0 0 6px rgba(184,134,11,.28);}}
  @media (prefers-reduced-motion: reduce){ .picker.first .combo input{animation:none;} }
  .welcome{text-align:center; padding:26px 10px 22px;}
  .welcome .wt{font-family:'Shippori Mincho', serif; font-weight:800; font-size:1.15rem; margin-bottom:8px;}
  .welcome p{font-size:.82rem; color:var(--ink-soft); line-height:1.8; margin:0 auto; max-width:30em; text-align:left;}

  .examcard{
    background:var(--tile); border:1px solid var(--tile-edge); border-radius:16px;
    box-shadow:var(--shadow); padding:26px 26px 22px; margin-bottom:46px;
  }
  .exam-head{display:flex; align-items:center; justify-content:space-between; flex-wrap:wrap; gap:10px; margin-bottom:14px;}
  .exam-name{font-family:'Shippori Mincho', serif; font-weight:800; font-size:1.5rem;}
  .exam-badges{display:flex; gap:8px; flex-wrap:wrap;}
  .rankbadge{display:inline-block;font-size:.72rem;padding:4px 12px;border-radius:20px;color:#fff;font-weight:700;letter-spacing:.05em;}
  .statusbadge{display:inline-block;font-size:.72rem;padding:4px 12px;border-radius:20px;font-weight:700;letter-spacing:.05em;border:1.5px solid;}
  .statusbadge.ready{color:var(--green-dragon);border-color:var(--green-dragon);background:#e9f5ec;}
  .statusbadge.on_track{color:#1f6b3a;border-color:#1f6b3a;background:#eef7f0;}
  .statusbadge.behind{color:var(--amber);border-color:var(--amber);background:#fbf1e3;}
  .statusbadge.reset_pending{color:var(--red);border-color:var(--red);background:#fbeceb;}
  .statusbadge.not_started{color:var(--ink-soft);border-color:var(--ink-soft);background:#f1efe8;}
  .statusbadge.unranked{color:var(--ink-soft);border-color:var(--ink-soft);background:#f1efe8;}

  .exam-msg{font-size:.82rem; color:var(--ink-soft); margin-bottom:18px; line-height:1.6;}

  .tilewall{display:flex; flex-wrap:wrap; gap:4px; margin-bottom:20px;}
  .tilewall .t{
    width:24px;height:32px;border-radius:4px;background:#efe6cd;border:1px solid var(--tile-edge);
    display:flex;align-items:center;justify-content:center;
    font-family:'Shippori Mincho', serif;font-weight:800;font-size:.95rem;color:#fff;
    box-shadow:inset 0 -3px 0 rgba(0,0,0,.18);
  }
  .tilewall .t.done{background:var(--green-dragon);border-color:var(--green-dragon);}
  .tilewall .t.p1{background:#1f6b3a;border-color:#174f2b;}
  .tilewall .t.p2{background:#2b74b8;border-color:#225d94;}
  .tilewall .t.p3{background:#b8780f;border-color:#93600c;}
  .tilewall .t.p4{background:#a8322c;border-color:#842520;}
  .tilewall .lbl{width:100%;font-size:.68rem;color:var(--ink-soft);margin-bottom:2px;display:flex;justify-content:space-between;flex-wrap:wrap;gap:4px 10px;}
  .placelegend{display:flex;gap:8px;flex-wrap:wrap;font-size:.66rem;color:var(--ink-soft);}
  .placelegend span{display:inline-flex;align-items:center;gap:3px;}
  .placelegend i{display:inline-block;width:10px;height:10px;border-radius:2px;}
  .placelegend b{font-weight:700;color:var(--ink);}

  .metrics{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:16px;}
  .metric{background:rgba(0,0,0,.03);border:1px solid var(--tile-edge);border-radius:10px;padding:14px 16px;}
  .metric .mlabel{font-size:.7rem;color:var(--ink-soft);letter-spacing:.05em;margin-bottom:6px;display:flex;justify-content:space-between;}
  .metric .mval{font-family:'Shippori Mincho', serif;font-size:1.3rem;font-weight:700;}
  .metric .mval .req{font-size:.78rem;font-weight:500;color:var(--ink-soft);margin-left:6px;}
  .mtag{font-size:.66rem;padding:2px 8px;border-radius:12px;font-weight:700;}
  .mtag.clear{background:#e9f5ec;color:var(--green-dragon);}
  .mtag.notclear{background:#fbeceb;color:var(--red);}

  .pacebox{background:rgba(184,134,11,.08); border:1px dashed var(--gold); border-radius:10px; padding:12px 16px; font-size:.8rem; color:#6b5410; line-height:1.6;}
  .pacebox.placebox{background:rgba(31,107,58,.07); border:1px dashed var(--green-dragon); color:#204d2f; margin-top:8px;}

  .simulator{margin-top:16px; padding:16px 18px; border-radius:12px; border:2px solid var(--tile-edge); background:rgba(0,0,0,.03); transition: background .2s ease, border-color .2s ease;}
  .simulator.safe{background:rgba(31,107,58,.12); border-color:var(--green-dragon);}
  .simulator.unsafe{background:rgba(168,50,44,.10); border-color:var(--red);}
  .simulator .simtitle{font-size:.76rem; letter-spacing:.05em; color:var(--ink-soft); margin-bottom:10px; font-weight:700;}
  .simulator .siminputs{display:grid; grid-template-columns:repeat(4,1fr); gap:10px; margin-bottom:12px;}
  .simulator .simcol{text-align:center;}
  .simulator .simcol label{display:block; font-size:.7rem; color:var(--ink-soft); margin-bottom:4px;}
  .simulator .simcol input{width:100%; text-align:center; font-family:'Shippori Mincho', serif; font-weight:700; font-size:1.1rem; padding:8px 4px; border-radius:8px; border:1px solid var(--tile-edge); background:var(--tile); color:var(--ink);}
  .simulator .simresult{font-size:.82rem; line-height:1.6; color:var(--ink);}
  .simulator .simresult b{font-family:'Shippori Mincho', serif; font-size:1.05rem;}
  .simulator .simreset{margin-top:8px; background:none; border:1px solid var(--ink-soft); color:var(--ink-soft); font-size:.7rem; padding:5px 12px; border-radius:14px; cursor:pointer;}

  /* ---- career + period two-panel ---- */
  .stats-two{
    margin-top:18px; padding-top:16px; border-top:1px dashed var(--tile-edge);
    display:grid; grid-template-columns:1fr 1fr; gap:14px;
  }
  .stats-panel{
    background:rgba(0,0,0,.03); border:1px solid var(--tile-edge); border-radius:10px;
    padding:14px 14px 10px;
  }
  .stats-panel .ptitle{
    font-family:'Shippori Mincho', serif; font-weight:700; font-size:.85rem;
    letter-spacing:.1em; color:var(--ink); text-align:center;
    padding-bottom:8px; margin-bottom:10px; border-bottom:1px solid var(--tile-edge);
  }
  .stats-panel .ptitle.career{color:var(--gold);}
  .stats-panel .ptitle.period{color:var(--green-dragon);}
  .stats-panel .pgrid{display:grid; grid-template-columns:repeat(2,1fr); gap:8px 12px; text-align:center;}
  .stats-panel .pgrid > div{font-size:.68rem; color:var(--ink-soft);}
  .stats-panel .pgrid b{display:block; font-family:'Shippori Mincho', serif; font-size:1rem; color:var(--ink); font-weight:700; margin-bottom:1px;}

  .controls{display:flex; gap:10px; flex-wrap:wrap; align-items:center; justify-content:space-between; margin-bottom:14px;}
  .sortbtns{display:flex; gap:6px; flex-wrap:wrap;}
  .sortbtns button{background:rgba(255,255,255,.08); border:1px solid rgba(255,255,255,.18); color:#eef3ee; font-family:'Zen Kaku Gothic New',sans-serif; font-size:.74rem; padding:7px 12px; border-radius:20px; cursor:pointer; letter-spacing:.05em; transition:.15s;}
  .sortbtns button:hover{background:rgba(255,255,255,.16);}
  .sortbtns button.active{background:var(--tile); color:var(--ink); border-color:var(--tile); font-weight:700;}
  .searchbox{background:rgba(255,255,255,.08); border:1px solid rgba(255,255,255,.18); border-radius:20px; padding:7px 14px;}
  .searchbox input{background:transparent; border:none; outline:none; color:#f6efdd; font-size:.8rem; width:150px; font-family:'Zen Kaku Gothic New',sans-serif;}
  .searchbox input::placeholder{color:rgba(255,255,255,.4);}

  .board{background:rgba(0,0,0,.15); border-radius:14px; padding:8px; border:1px solid rgba(255,255,255,.08);}
  .row{display:grid; grid-template-columns:44px 1.3fr 78px 66px 60px 74px; align-items:center; gap:6px; background:var(--tile); border:1px solid var(--tile-edge); border-radius:8px; padding:11px 12px; margin-bottom:6px; font-size:.82rem; cursor:pointer; transition:transform .12s ease;}
  .row:hover{transform:translateX(2px);}
  .row.selected{outline:2px solid var(--gold);}
  .row.head{background:transparent;border:none;color:rgba(255,255,255,.5);font-size:.64rem;letter-spacing:.08em;padding:2px 12px 8px;margin-bottom:0;cursor:default;}
  .row.head:hover{transform:none;}
  .row.head .num{color:inherit;}
  .row .rk{font-family:'Shippori Mincho', serif;font-weight:800;font-size:.95rem;color:var(--ink-soft);text-align:center;}
  .row .nm{font-weight:700;color:var(--ink);overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
  .row .badge{justify-self:start;font-size:.66rem;padding:3px 9px;border-radius:20px;color:#fff;text-align:center;white-space:nowrap;}
  .row .num{text-align:right;color:var(--ink);font-variant-numeric:tabular-nums;}
  .row .st{justify-self:end;font-size:.62rem;padding:3px 8px;border-radius:12px;font-weight:700;text-align:center;white-space:nowrap;}
  .row .st.ready{background:#e9f5ec;color:var(--green-dragon);}
  .row .st.on_track{background:#eef7f0;color:#1f6b3a;}
  .row .st.behind{background:#fbf1e3;color:var(--amber);}
  .row .st.reset_pending{background:#fbeceb;color:var(--red);}
  .row .st.not_started{background:#f1efe8;color:var(--ink-soft);}
  .row .st.unranked{background:#f1efe8;color:var(--ink-soft);}

  .empty{text-align:center;color:rgba(255,255,255,.5);padding:40px;font-size:.85rem;}
  footer{text-align:center;color:rgba(255,255,255,.35);font-size:.7rem;margin-top:40px;letter-spacing:.05em;}

  /* ---- copy link ---- */
  .linkbtn{
    background:rgba(255,255,255,.08); border:1px solid rgba(255,255,255,.25); color:#eef3ee;
    font-family:'Zen Kaku Gothic New',sans-serif; font-size:.78rem; padding:9px 14px;
    border-radius:20px; cursor:pointer; letter-spacing:.03em; touch-action:manipulation;
  }
  .linkbtn.done{background:var(--tile); color:var(--green-dragon); border-color:var(--tile); font-weight:700;}

  /* ---- news (今回の更新) ---- */
  .news{
    background:rgba(255,255,255,.06); border:1px solid rgba(255,255,255,.14); border-radius:12px;
    padding:14px 16px 6px; margin:0 0 20px; color:#eef3ee;
  }
  .news .ntitle{font-family:'Shippori Mincho', serif; font-weight:700; letter-spacing:.12em; font-size:.92rem; color:#f6efdd; margin-bottom:10px;}
  .news .ntitle small{font-family:'Zen Kaku Gothic New',sans-serif; font-weight:400; letter-spacing:.03em; font-size:.68rem; color:rgba(255,255,255,.5); margin-left:8px;}
  .news .ngroup{display:flex; gap:10px; align-items:flex-start; margin-bottom:10px;}
  .news .nlabel{flex:0 0 8.4em; font-size:.72rem; font-weight:700; padding-top:5px; color:rgba(255,255,255,.75);}
  .news .nlabel .ico{display:inline-block; width:1.3em; text-align:center;}
  .news .nchips{display:flex; flex-wrap:wrap; gap:6px;}
  .news .nchip{
    background:var(--tile); color:var(--ink); border:1px solid var(--tile-edge); border-radius:16px;
    padding:4px 10px; font-size:.76rem; font-weight:700; cursor:pointer; font-family:'Zen Kaku Gothic New',sans-serif;
  }
  .news .nchip small{font-weight:500; color:var(--ink-soft); margin-left:4px;}
  .news .nnone{font-size:.74rem; color:rgba(255,255,255,.45); padding-top:5px;}

  /* ---- status filter ---- */
  .filterbtns{display:flex; gap:6px; flex-wrap:wrap; justify-content:center; margin:0 0 12px;}
  .filterbtns button{
    background:transparent; border:1px dashed rgba(255,255,255,.25); color:rgba(255,255,255,.75);
    font-family:'Zen Kaku Gothic New',sans-serif; font-size:.72rem; padding:6px 11px; border-radius:16px; cursor:pointer;
  }
  .filterbtns button b{margin-left:4px; font-weight:700;}
  .filterbtns button.active{background:var(--tile); color:var(--ink); border:1px solid var(--tile); font-weight:700;}

  /* ---- simulator: 収支判定 ---- */
  .simjudge{display:grid; grid-template-columns:1fr 1fr; gap:8px; margin-bottom:10px;}
  .simjudge > div{background:rgba(255,255,255,.55); border:1px solid var(--tile-edge); border-radius:8px; padding:8px 10px; font-size:.72rem; color:var(--ink-soft);}
  .simjudge b{display:block; font-family:'Shippori Mincho', serif; font-size:1.15rem; color:var(--ink);}
  .simjudge .ok{color:var(--green-dragon); font-weight:700;}
  .simjudge .ng{color:var(--red); font-weight:700;}
  .simverdict{font-size:.86rem; font-weight:700; margin-bottom:6px;}
  .simnote{font-size:.68rem; color:var(--ink-soft); line-height:1.6;}

  /* ---- trend charts ---- */
  .trends{margin-top:18px; padding-top:16px; border-top:1px dashed var(--tile-edge); display:grid; grid-template-columns:1fr; gap:14px;}
  .trend{background:rgba(0,0,0,.03); border:1px solid var(--tile-edge); border-radius:10px; padding:12px 12px 8px; position:relative;}
  .trend .ttitle{font-family:'Shippori Mincho', serif; font-weight:700; font-size:.85rem; color:var(--ink); margin-bottom:2px;}
  .trend .tsub{font-size:.66rem; color:var(--ink-soft); margin-bottom:6px;}
  .trend svg{display:block; width:100%; height:150px; overflow:visible; touch-action:pan-y;}
  .trend svg text{font-family:'Zen Kaku Gothic New',sans-serif; font-size:10px; fill:var(--ink-soft);}
  .trend .tip{
    position:absolute; pointer-events:none; background:var(--ink); color:#f6efdd; font-size:.7rem;
    padding:4px 8px; border-radius:6px; white-space:nowrap; transform:translate(-50%,-100%); display:none; z-index:2;
  }
  .trend .tempty{font-size:.74rem; color:var(--ink-soft); padding:30px 0; text-align:center;}

  /* ---- simulator stepper (+/-) ---- */
  .stepper{display:flex; align-items:stretch; gap:4px;}
  .stepper button{
    flex:0 0 30px; border:1px solid var(--tile-edge); background:#efe6cd; color:var(--ink);
    border-radius:8px; font-size:1.1rem; font-weight:700; line-height:1; cursor:pointer;
    padding:0; touch-action:manipulation;
  }
  .stepper button:active{background:var(--tile-edge);}
  .stepper input{min-width:0;}
  .stepper input::-webkit-outer-spin-button, .stepper input::-webkit-inner-spin-button{-webkit-appearance:none; margin:0;}
  .stepper input[type=number]{-moz-appearance:textfield;}

  html{-webkit-text-size-adjust:100%;}
  body{-webkit-tap-highlight-color:transparent;}
  @media (hover:none){
    .row:hover{transform:none;}
    .sortbtns button:hover{background:rgba(255,255,255,.08);}
  }

  /* ---- tablet / small laptop ---- */
  @media (max-width:720px){
    .statbar{grid-template-columns:repeat(2,1fr);}
    .stats-two{grid-template-columns:1fr;}
    .trends{grid-template-columns:1fr;}
  }

  /* ---- smartphone ---- */
  @media (max-width:600px){
    body{padding-bottom:48px;}
    .wrap{padding:12px 14px 0;}
    header.plaque{padding:18px 0 20px;}
    .kanban{padding:14px 26px;}
    .kanban h1{font-size:1.9rem; letter-spacing:.1em; white-space:nowrap;}
    .kanban .sub{font-size:.74rem; letter-spacing:.14em;}
    .updated{margin-top:10px;}

    .statbar{gap:8px; margin-bottom:28px;}
    .stat{padding:10px 6px;}
    .stat .n{font-size:1.3rem;}

    h2.section-title{margin-bottom:14px;}
    .picker{margin-bottom:14px;}
    .combo{max-width:none;}
    .linkbtn{width:100%;}
    .news{padding:12px 12px 4px; margin-bottom:16px;}
    .news .ngroup{flex-direction:column; gap:6px;}
    .news .nlabel{flex:none; padding-top:0;}
    .filterbtns{flex-wrap:nowrap; overflow-x:auto; justify-content:flex-start; scrollbar-width:none; -webkit-overflow-scrolling:touch;}
    .filterbtns::-webkit-scrollbar{display:none;}
    .filterbtns button{flex:0 0 auto;}
    .trend svg{height:140px;}

    .examcard{padding:18px 14px 16px; border-radius:14px; margin-bottom:34px;}
    .exam-head{margin-bottom:10px; gap:8px;}
    .exam-name{font-size:1.3rem;}
    .exam-msg{margin-bottom:14px;}

    .tilewall{gap:3px; margin-bottom:16px;}
    .tilewall .t{width:21px; height:28px; font-size:.85rem;}

    .metrics{gap:8px; margin-bottom:12px;}
    .metric{padding:10px 10px;}
    .metric .mlabel{flex-direction:column; align-items:flex-start; gap:4px;}
    .metric .mval{font-size:1.25rem;}
    .metric .mval .req{display:block; margin:2px 0 0;}

    .pacebox{padding:10px 12px;}
    .simulator{padding:14px 12px;}
    .simulator .siminputs{grid-template-columns:repeat(2,1fr); gap:10px 12px;}
    .stepper button{flex-basis:40px; font-size:1.25rem;}
    .simulator .simcol input{font-size:1.2rem; padding:9px 2px;}
    .simulator .simreset{font-size:.76rem; padding:7px 14px;}

    .stats-two{gap:10px; margin-top:14px; padding-top:14px;}
    .stats-panel{padding:10px 10px 8px;}
    .stats-panel .pgrid{grid-template-columns:repeat(4,1fr); gap:6px 4px;}
    .stats-panel .pgrid > div{font-size:.6rem; line-height:1.3;}
    .stats-panel .pgrid b{font-size:.95rem;}

    .controls{flex-direction:column; align-items:stretch; gap:10px;}
    .sortbtns{flex-wrap:nowrap; overflow-x:auto; -webkit-overflow-scrolling:touch; scrollbar-width:none; padding-bottom:2px;}
    .sortbtns::-webkit-scrollbar{display:none;}
    .sortbtns button{flex:0 0 auto; padding:8px 13px;}
    .searchbox{padding:8px 14px;}
    .searchbox input{width:100%; font-size:16px;}

    .board{padding:6px;}
    .row.head{display:none;}
    .row{
      grid-template-columns:30px auto auto 1fr;
      grid-template-areas:"rk nm nm badge" "rk place games st";
      gap:4px 10px; padding:10px 12px; font-size:.86rem;
    }
    .row .rk{grid-area:rk;}
    .row .nm{grid-area:nm;}
    .row > div:nth-child(3){grid-area:badge; justify-self:end;}
    .row .col-place{grid-area:place;}
    .row .col-games{grid-area:games;}
    .row .col-status{grid-area:st;}
    .row .col-place, .row .col-games{text-align:left; font-size:.72rem; color:var(--ink-soft);}
    .row .col-place::before{content:"平均着順 ";}
    .row .col-games::before{content:"半荘 ";}
    .row .col-place b, .row .col-games b{color:var(--ink); font-size:.82rem;}
  }
</style>
</head>
<body>
<div class="wrap">

  <header class="plaque">
    <div class="kanban">
      <h1>らいらいず</h1>
      <div class="sub">個人 成績 ・ 段位 進捗</div>
    </div>
    <div class="updated">__GENERATED_AT__</div>
  </header>

  <h2 class="section-title">昇 段 診 断<span class="asof">__DATE_ONLY__</span></h2>
  <div class="picker">
    <div class="combo">
      <input id="playerSearch" type="text" placeholder="自分の名前を入力して検索" autocomplete="off" autocapitalize="off" spellcheck="false"
        role="combobox" aria-autocomplete="list" aria-expanded="false" aria-controls="playerList" aria-label="打ち手を名前で検索">
      <ul id="playerList" role="listbox" hidden></ul>
    </div>
    <button type="button" class="linkbtn" id="copyLink" hidden>この人のリンクをコピー</button>
  </div>
  <div class="examcard" id="examcard"></div>

  <div id="news"></div>

  <div class="statbar" id="statbar"></div>



  <h2 class="section-title">らいらいずメンバー一覧</h2>
  <div class="filterbtns" id="filterbtns" aria-label="審査状況で絞り込み"></div>
  <div class="controls">
    <div class="sortbtns" id="sortbtns"></div>
    <div class="searchbox"><input id="search" type="text" placeholder="名前で検索..."></div>
  </div>
  <div class="board">
    <div class="row head">
      <div>順</div>
      <div>名前</div>
      <div>段位</div>
      <div class="num col-place">平均着順</div>
      <div class="num col-games">半荘</div>
      <div class="col-status" style="text-align:right;">審査状況</div>
    </div>
    <div id="rows"></div>
  </div>

  <footer>らいらいず 個人成績管理表より自動生成 &mdash; 段位基準はサークル既定値に基づく</footer>
</div>

<script id="rawData" type="application/json">__DATA_JSON__</script>
<script>
const DATA = JSON.parse(document.getElementById('rawData').textContent);
const RANK_ORDER = DATA.rank_order;
const players = DATA.players;

function rankTier(rank){
  if(!rank) return '#8a6a3d';
  if(rank.includes('九段')) return '#a8322c';
  if(rank.includes('八段')||rank.includes('七段')) return '#b8860b';
  if(rank.includes('六段')||rank.includes('五段')) return '#8a6a20';
  if(rank.includes('段')) return '#1f6b3a';
  if(rank.includes('1級')||rank.includes('2級')) return '#5b6b7c';
  if(rank.includes('級')) return '#8a8a8a';
  return '#8a6a3d';
}

function fmt(v, digits){
  if(v===null || v===undefined || isNaN(v)) return '—';
  return v.toFixed(digits);
}
function fmtPts(v){
  if(v===null||v===undefined) return '—';
  const sign = v>0?'+':'';
  return sign + v.toFixed(1);
}

const STATUS_LABEL = {
  ready:'合格ライン到達', on_track:'基準クリア中', behind:'追い上げ中',
  reset_pending:'要再挑戦', not_started:'審査開始前', unranked:'未認定',
};

let selectedKey = null;
let activePlayers = players.filter(p=>p.rank_value>=0);

function renderStatbar(){
  const n = activePlayers.length;
  const totalGames = activePlayers.reduce((a,p)=>a+(p.career.games||0),0);
  const readyCount = activePlayers.filter(p=>p.exam && p.exam.status==='ready').length;
  const onTrackCount = activePlayers.filter(p=>p.exam && p.exam.status==='on_track').length;
  document.getElementById('statbar').innerHTML = `
    <div class="stat"><div class="n">${n}</div><div class="l">在籍人数</div></div>
    <div class="stat"><div class="n">${totalGames}</div><div class="l">累計半荘数（全員）</div></div>
    <div class="stat"><div class="n">${readyCount}</div><div class="l">合格ライン到達中</div></div>
    <div class="stat"><div class="n">${onTrackCount}</div><div class="l">基準クリア継続中</div></div>
  `;
}

// ---- 自分専用リンク（?p=キー）と前回見た人の記憶 ----
const LAST_KEY_STORE = 'rairise:lastKey';
function findPlayer(q){
  if(!q) return null;
  return players.find(p=>p.key===q) || players.find(p=>p.name===q) || null;
}
function storeGet(k){ try{ return localStorage.getItem(k); }catch(e){ return null; } }
function storeSet(k,v){ try{ localStorage.setItem(k,v); }catch(e){} }
function playerUrl(key){
  const u = new URL(location.href);
  u.search = ''; u.hash = '';
  u.searchParams.set('p', key);
  return u.toString();
}
function selectPlayer(key, opts){
  opts = opts || {};
  selectedKey = key;
  const pl = players.find(x=>x.key===key);
  document.getElementById('playerSearch').value = pl ? pl.name : '';
  document.querySelector('.picker').classList.remove('first');
  storeSet(LAST_KEY_STORE, key);
  try{ history.replaceState(null, '', playerUrl(key)); }catch(e){}
  renderExamCard();
  highlightRow();
  const btn = document.getElementById('copyLink');
  btn.hidden = false; btn.classList.remove('done'); btn.textContent = 'この人のリンクをコピー';
  if(opts.scroll) document.getElementById('examcard').scrollIntoView({behavior:'smooth', block:'start'});
}

// ひらがな/カタカナ・全角/半角・大文字小文字を区別せずに検索する
function normName(t){
  return (t||'').normalize('NFKC').toLowerCase()
    .replace(/[ァ-ヶ]/g, c=>String.fromCharCode(c.charCodeAt(0)-0x60))
    .replace(/[\s・☆★]/g,'');
}

function renderPicker(){
  const input = document.getElementById('playerSearch');
  const list = document.getElementById('playerList');
  const sorted = [...players].sort((a,b)=> b.rank_value - a.rank_value || a.name.localeCompare(b.name,'ja'));
  let matches = [], active = -1;

  function open(){
    const q = normName(input.value);
    matches = sorted.filter(p => !q || normName(p.name).includes(q) || normName(p.key).includes(q));
    // 前方一致を先に
    if(q) matches.sort((a,b)=> (normName(b.name).startsWith(q)) - (normName(a.name).startsWith(q)));
    active = matches.length ? 0 : -1;
    list.innerHTML = matches.length
      ? matches.map((p,i)=>`<li role="option" id="opt-${i}" data-key="${p.key}" aria-selected="${i===active}"><span>${p.name}</span><span class="badge" style="background:${rankTier(p.rank)}">${p.rank||'未認定'}</span></li>`).join('')
      : '<li class="none">該当する名前がありません</li>';
    list.hidden = false;
    input.setAttribute('aria-expanded','true');
  }
  function close(restore){
    list.hidden = true;
    input.setAttribute('aria-expanded','false');
    if(restore){ const p = players.find(x=>x.key===selectedKey); input.value = p ? p.name : ''; }
  }
  function move(d){
    if(!matches.length) return;
    active = (active + d + matches.length) % matches.length;
    list.querySelectorAll('li').forEach((li,i)=> li.setAttribute('aria-selected', i===active));
    const li = list.querySelector(`#opt-${active}`); if(li){ li.scrollIntoView({block:'nearest'}); input.setAttribute('aria-activedescendant', li.id); }
  }
  function choose(key){ close(false); input.blur(); selectPlayer(key, {scroll:true}); }

  input.addEventListener('focus', ()=>{ input.select(); open(); });
  input.addEventListener('input', open);
  input.addEventListener('keydown', e=>{
    if(e.key==='ArrowDown'){ e.preventDefault(); if(list.hidden) open(); else move(1); }
    else if(e.key==='ArrowUp'){ e.preventDefault(); move(-1); }
    else if(e.key==='Enter'){ e.preventDefault(); if(active>=0 && matches[active]) choose(matches[active].key); }
    else if(e.key==='Escape'){ close(true); input.blur(); }
  });
  // スクロール中のタップ誤選択を避けるため click で確定（mousedown で入力のblurを防ぐ）
  list.addEventListener('mousedown', e=> e.preventDefault());
  list.addEventListener('click', e=>{ const li = e.target.closest('li[data-key]'); if(li) choose(li.dataset.key); });
  document.addEventListener('pointerdown', e=>{ if(!e.target.closest('.combo')) close(true); });

  let fromUrl = null;
  try{ fromUrl = findPlayer(new URLSearchParams(location.search).get('p')); }catch(e){}
  const remembered = findPlayer(storeGet(LAST_KEY_STORE));
  const initial = fromUrl || remembered;
  if(initial){
    selectedKey = initial.key;
    input.value = initial.name;
    document.getElementById('copyLink').hidden = false;
    if(fromUrl) storeSet(LAST_KEY_STORE, fromUrl.key);
  } else {
    selectedKey = null;
    document.querySelector('.picker').classList.add('first');
  }

  document.getElementById('copyLink').addEventListener('click', async (ev)=>{
    const btn = ev.currentTarget;
    const url = playerUrl(selectedKey);
    let ok = false;
    try{ await navigator.clipboard.writeText(url); ok = true; }catch(e){
      try{
        const ta = document.createElement('textarea'); ta.value = url; ta.setAttribute('readonly','');
        ta.style.position='fixed'; ta.style.opacity='0'; document.body.appendChild(ta); ta.select();
        ok = document.execCommand('copy'); ta.remove();
      }catch(e2){}
    }
    btn.classList.add('done');
    btn.textContent = ok ? 'コピーしました（ブックマークにも使えます）' : url;
  });
}

// ---- 今回の更新で動いたこと ----
function renderNews(){
  const ch = DATA.changes;
  const el = document.getElementById('news');
  if(!ch){ el.innerHTML=''; return; }
  const byKey = Object.fromEntries(players.map(p=>[p.key,p]));
  const chip = (key, sub) => byKey[key] ? `<button type="button" class="nchip" data-key="${key}">${byKey[key].name}${sub?`<small>${sub}</small>`:''}</button>` : '';
  const group = (ico, label, items) => `
    <div class="ngroup"><div class="nlabel"><span class="ico">${ico}</span>${label}</div>
      <div class="nchips">${items.length ? items.join('') : '<span class="nnone">該当なし</span>'}</div></div>`;
  const promoted = (ch.promoted||[]).map(x=>chip(x.key, `${x.from}→${x.to}`));
  const ready = (ch.ready||[]).map(k=>chip(k, byKey[k] && byKey[k].exam ? byKey[k].exam.rank : ''));
  const reach = (ch.reach||[]).map(k=>chip(k, byKey[k] && byKey[k].exam ? `あと${byKey[k].exam.remaining_games}半荘` : ''));
  const fresh = (ch.new||[]).map(k=>chip(k, ''));
  let html = `<div class="news"><div class="ntitle">今回の更新${ch.since?`<small>${ch.since} からの変化</small>`:''}</div>`;
  if(ch.since) html += group('▲', '昇段した人', promoted);
  html += group('◎', '合格ライン到達', ready);
  html += group('◆', `昇段リーチ`, reach);
  if(fresh.length) html += group('＋', '新メンバー', fresh);
  html += `</div>`;
  el.innerHTML = html;
  el.querySelectorAll('.nchip').forEach(b=>b.addEventListener('click', ()=> selectPlayer(b.dataset.key, {scroll:true})));
}

const PLACE_LABEL = {1:'トップ', 2:'二着', 3:'三着', 4:'ラス'};
const PLACE_COLOR = {1:'#1f6b3a',2:'#2b74b8',3:'#b8780f',4:'#a8322c'};

function placeExampleHtml(ex){
  const remaining = ex.remaining_games;
  const placeSum = ex.avg_place!==null ? ex.place_sum : 0;
  const budget = ex.required_place * ex.required_games - placeSum;

  if(budget < remaining * 1 - 1e-9){
    return `<div class="pacebox placebox">着順基準：残り${remaining}半荘では、全対局トップを取っても平均着順の基準到達は数値上厳しい状況です。早めのペースアップが必要です。</div>`;
  }
  if(budget >= remaining * 4 - 1e-9){
    return `<div class="pacebox placebox">着順基準：残り${remaining}半荘は着順を問わず基準内に収まります（ポイント基準のみ意識すればOK）。</div>`;
  }
  let lower = Math.floor(budget / remaining);
  lower = Math.max(1, Math.min(3, lower));
  const higher = lower + 1;
  let nHigher = Math.floor(budget - remaining * lower + 1e-9);
  nHigher = Math.max(0, Math.min(remaining, nHigher));
  const nLower = remaining - nHigher;
  const counts = {1:0,2:0,3:0,4:0};
  counts[lower] += nLower;
  counts[higher] += nHigher;
  const parts = [1,2,3,4].map(k => `${PLACE_LABEL[k]}${counts[k]}回`).join('　');
  return `<div class="pacebox placebox">着順基準：残り${remaining}半荘の一例として、<b>${parts}</b> のペースならセーフ（平均着順${ex.required_place}以内を維持できます）。</div>`;
}

function simAssumedScores(p){
  // 着順ごとの想定素点：本人の過去平均（3回以上あれば）→ なければサークル全体の平均
  const mine = (p.career && p.career.place_avg_score) || {};
  const circle = DATA.circle_place_avg || {};
  const out = {}, src = {};
  [1,2,3,4].forEach(k=>{
    const m = mine[k] || mine[String(k)];
    if(m && m.n >= 3 && m.avg !== null){ out[k] = m.avg; src[k] = 'mine'; }
    else { out[k] = circle[k] ?? circle[String(k)] ?? 0; src[k] = 'circle'; }
  });
  return {score: out, src};
}

function simulatorHtml(ex){
  if(!ex.remaining_games || ex.remaining_games <= 0) return '';
  return `
  <div class="simulator" id="simulator">
    <div class="simtitle">昇段シミュレーター　（残り ${ex.remaining_games} 半荘のうち、想定回数を入力）</div>
    <div class="siminputs">
      ${[1,2,3,4].map(k=>`
      <div class="simcol"><label for="sim-${k}">${PLACE_LABEL[k]}</label>
        <div class="stepper">
          <button type="button" class="step" data-k="${k}" data-d="-1" aria-label="${PLACE_LABEL[k]}を1減らす">−</button>
          <input type="number" min="0" inputmode="numeric" pattern="[0-9]*" id="sim-${k}" value="0">
          <button type="button" class="step" data-k="${k}" data-d="1" aria-label="${PLACE_LABEL[k]}を1増やす">＋</button>
        </div>
      </div>`).join('')}
    </div>
    <div class="simresult" id="simresult"></div>
    <button class="simreset" id="simreset">入力をリセット</button>
  </div>`;
}

function bindSimulator(ex, p){
  if(!ex.remaining_games || ex.remaining_games <= 0) return;
  const ids = [1,2,3,4].map(k=>document.getElementById('sim-'+k));
  const box = document.getElementById('simulator');
  const resultEl = document.getElementById('simresult');
  const cap = ex.remaining_games;
  const assumed = simAssumedScores(p);

  ids.forEach(el => el.setAttribute('max', cap));

  function clampField(changedEl){
    let counts = ids.map(el => Math.max(0, parseInt(el.value)||0));
    let total = counts.reduce((a,b)=>a+b,0);
    if(total > cap){
      const idx = ids.indexOf(changedEl);
      counts[idx] = Math.max(0, counts[idx] - (total - cap));
      changedEl.value = counts[idx];
    } else {
      changedEl.value = counts[ids.indexOf(changedEl)];
    }
    const newCounts = ids.map(el => Math.max(0, parseInt(el.value)||0));
    const newTotal = newCounts.reduce((a,b)=>a+b,0);
    ids.forEach((el,i) => el.setAttribute('max', cap - (newTotal - newCounts[i])));
  }

  function update(){
    const counts = ids.map(el => Math.max(0, parseInt(el.value)||0));
    const entered = counts.reduce((a,b)=>a+b,0);
    const combinedGames = (ex.played_games||0) + entered;

    // 着順
    const enteredPlaceSum = counts[0]*1 + counts[1]*2 + counts[2]*3 + counts[3]*4;
    const combinedPlace = combinedGames>0 ? ((ex.place_sum||0) + enteredPlaceSum)/combinedGames : null;
    const placeOk = combinedGames>0 && combinedPlace <= ex.required_place;

    // 収支（着順ごとの想定素点で推定）
    const enteredScore = counts.reduce((a,c,i)=> a + c*assumed.score[i+1], 0);
    const combinedScore = combinedGames>0 ? ((ex.total_current||0) + enteredScore)/combinedGames : null;
    const scoreOk = combinedGames>0 && combinedScore >= ex.required_avg;

    const both = placeOk && scoreOk;
    box.classList.remove('safe','unsafe');
    box.classList.add(both ? 'safe':'unsafe');

    const full = entered === cap;
    let verdict;
    if(entered === 0) verdict = '回数を入力すると、昇段できるかを判定します。';
    else if(both) verdict = full ? '✓ この想定なら、着順・収支とも基準クリアで昇段ラインです。' : '✓ ここまでの想定では、着順・収支とも基準クリアです。';
    else if(placeOk) verdict = '△ 着順はクリアですが、収支が基準に届きません。';
    else if(scoreOk) verdict = '△ 収支はクリアですが、着順が基準に届きません。';
    else verdict = '× 着順・収支とも基準に届きません。';

    const assumedTxt = [1,2,3,4].map(k=>`${PLACE_LABEL[k]} ${fmtPts(assumed.score[k])}${assumed.src[k]==='circle'?'*':''}`).join(' / ');
    const hasCircle = [1,2,3,4].some(k=>assumed.src[k]==='circle');

    resultEl.innerHTML = `
      <div class="simverdict" style="color:${entered===0?'var(--ink-soft)':(both?'var(--green-dragon)':'var(--red)')}">${verdict}</div>
      <div class="simjudge">
        <div>平均着順（${ex.required_place}以内）<b>${combinedGames>0? combinedPlace.toFixed(3):'—'}</b><span class="${placeOk?'ok':'ng'}">${placeOk?'クリア':'未達'}</span></div>
        <div>平均収支（${ex.required_avg}以上）<b>${combinedGames>0? combinedScore.toFixed(1):'—'}</b><span class="${scoreOk?'ok':'ng'}">${scoreOk?'クリア':'未達'}</span></div>
      </div>
      <div class="simnote">入力合計 ${entered} / 残り${cap}半荘${full?'（上限まで入力済み）':`（あと${cap-entered}半荘分入力できます）`}<br>
      収支は1半荘あたりの想定素点で推定：${assumedTxt}<br>
      （${p.name}さんの過去の着順別平均${hasCircle?'。*は過去データが少ないためサークル全体の平均':''}）</div>`;
  }

  ids.forEach(el => el.addEventListener('input', ()=>{ clampField(el); update(); }));
  box.querySelectorAll('button.step').forEach(btn => btn.addEventListener('click', ()=>{
    const el = ids[parseInt(btn.dataset.k)-1];
    el.value = Math.max(0, (parseInt(el.value)||0) + parseInt(btn.dataset.d));
    clampField(el); update();
  }));
  document.getElementById('simreset').addEventListener('click', ()=>{
    ids.forEach(el => { el.value = 0; el.setAttribute('max', cap); });
    update();
  });
  update();
}

// ---- 成績推移グラフ（1系列ずつ・軸は1本） ----
function trendsHtml(p){
  const n = (p.career && p.career.places || []).length;
  if(n < 2) return `<div class="trends"><div class="trend"><div class="ttitle">成績の推移</div><div class="tempty">対局データが2半荘以上になると表示されます</div></div></div>`;
  const win = Math.min(10, n);
  return `
  <div class="trends">
    <div class="trend" id="trend-place">
      <div class="ttitle">平均着順の推移</div>
      <div class="tsub">直近${win}半荘の移動平均（1半荘目 → 最新${n}半荘）。上ほど好成績、網かけは現在の審査期間</div>
      <svg role="img" aria-label="平均着順の推移"></svg><div class="tip"></div>
    </div>
  </div>`;
}

function niceTicks(min, max, count){
  const span = max - min || 1;
  const step0 = span / count;
  const mag = Math.pow(10, Math.floor(Math.log10(step0)));
  const step = [1,2,2.5,5,10].map(m=>m*mag).find(s=>span/s <= count) || 10*mag;
  const t = []; for(let v=Math.ceil(min/step)*step; v<=max+1e-9; v+=step) t.push(+v.toFixed(6));
  return t;
}

function drawLine(container, opts){
  const svg = container.querySelector('svg');
  const tip = container.querySelector('.tip');
  const W = Math.max(svg.clientWidth || 300, 200), H = svg.clientHeight || 150;
  const m = {l:44, r:10, t:10, b:22};
  const vals = opts.values, n = vals.length;
  let lo = opts.min ?? Math.min(...vals.filter(v=>v!==null)), hi = opts.max ?? Math.max(...vals.filter(v=>v!==null));
  if(opts.ref !== undefined){ lo = Math.min(lo, opts.ref); hi = Math.max(hi, opts.ref); }
  if(opts.zero){ lo = Math.min(lo, 0); hi = Math.max(hi, 0); }
  if(hi === lo){ hi += 1; lo -= 1; }
  const pad = (hi-lo)*0.06; if(opts.min===undefined) lo -= pad; if(opts.max===undefined) hi += pad;
  const x = i => m.l + (n===1?0:(i/(n-1))*(W-m.l-m.r));
  const y = v => opts.invert ? m.t + ((v-lo)/(hi-lo))*(H-m.t-m.b) : H-m.b - ((v-lo)/(hi-lo))*(H-m.t-m.b);
  const NS = 'http://www.w3.org/2000/svg';
  let g = '';
  // 審査期間の網かけ
  if(opts.shadeFrom !== undefined && opts.shadeFrom < n){
    const x0 = x(Math.max(opts.shadeFrom-0.5,0)), x1 = x(n-1);
    g += `<rect x="${x0}" y="${m.t}" width="${Math.max(x1-x0,2)}" height="${H-m.t-m.b}" fill="#b8860b" opacity=".08"/>`;
  }
  // グリッドと目盛り（細く控えめに）
  const ticks = opts.ticks || niceTicks(lo, hi, 4);
  ticks.forEach(v=>{
    if(v<lo-1e-9||v>hi+1e-9) return;
    g += `<line x1="${m.l}" x2="${W-m.r}" y1="${y(v)}" y2="${y(v)}" stroke="#d9cba5" stroke-width="1"/>`;
    g += `<text x="${m.l-6}" y="${y(v)+3}" text-anchor="end">${opts.tickFmt(v)}</text>`;
  });
  // 基準線
  if(opts.ref !== undefined){
    g += `<line x1="${m.l}" x2="${W-m.r}" y1="${y(opts.ref)}" y2="${y(opts.ref)}" stroke="#5b6357" stroke-width="1"/>`;

  }
  // x軸ラベル
  g += `<text x="${m.l}" y="${H-6}" text-anchor="start">1</text><text x="${W-m.r}" y="${H-6}" text-anchor="end">${n}半荘</text>`;
  // 線
  let d=''; let started=false;
  vals.forEach((v,i)=>{ if(v===null){started=false;return;} d += (started?'L':'M') + x(i).toFixed(1) + ' ' + y(v).toFixed(1); started=true; });
  g += `<path d="${d}" fill="none" stroke="${opts.color}" stroke-width="2" stroke-linejoin="round" stroke-linecap="round"/>`;
  if(opts.ref !== undefined){
    g += `<text x="${m.l+4}" y="${y(opts.ref)-5}" text-anchor="start" style="fill:#20281f;font-weight:700;paint-order:stroke;stroke:#f1ead6;stroke-width:5px;stroke-linejoin:round">${opts.refLabel}</text>`;
  }
  // 最新値の点＋ラベル
  const li = n-1;
  if(vals[li]!==null){
    g += `<circle cx="${x(li)}" cy="${y(vals[li])}" r="4" fill="${opts.color}" stroke="#f6efdd" stroke-width="2"/>`;
  }
  // ホバー用
  g += `<line class="cross" x1="0" x2="0" y1="${m.t}" y2="${H-m.b}" stroke="#20281f" stroke-width="1" opacity="0"/>`;
  g += `<circle class="hdot" r="4" fill="${opts.color}" stroke="#f6efdd" stroke-width="2" opacity="0"/>`;
  g += `<rect class="hit" x="${m.l}" y="0" width="${W-m.l-m.r}" height="${H}" fill="transparent"/>`;
  svg.setAttribute('viewBox', `0 0 ${W} ${H}`);
  svg.innerHTML = g;
  const cross = svg.querySelector('.cross'), hdot = svg.querySelector('.hdot'), hit = svg.querySelector('.hit');
  function show(evt){
    const r = svg.getBoundingClientRect();
    const px = (evt.clientX - r.left) * (W / r.width);
    const i = Math.max(0, Math.min(n-1, Math.round((px - m.l)/((W-m.l-m.r)/Math.max(n-1,1)))));
    if(vals[i]===null) return;
    cross.setAttribute('x1', x(i)); cross.setAttribute('x2', x(i)); cross.setAttribute('opacity', .35);
    hdot.setAttribute('cx', x(i)); hdot.setAttribute('cy', y(vals[i])); hdot.setAttribute('opacity', 1);
    const cr = container.getBoundingClientRect();
    tip.textContent = opts.tipFmt(i, vals[i]);
    tip.style.display='block';
    const tw = tip.offsetWidth;
    let left = r.left - cr.left + x(i)*(r.width/W);
    left = Math.max(tw/2 + 4, Math.min(cr.width - tw/2 - 4, left));
    tip.style.left = left + 'px';
    tip.style.top = (r.top - cr.top + y(vals[i])*(r.height/H) - 10) + 'px';
  }
  function hide(){ cross.setAttribute('opacity',0); hdot.setAttribute('opacity',0); tip.style.display='none'; }
  hit.addEventListener('pointermove', show);
  hit.addEventListener('pointerdown', show);
  hit.addEventListener('pointerleave', hide);
}

function renderTrends(p, ex){
  const places = (p.career && p.career.places) || [];
  const n = places.length;
  if(n < 2) return;
  const shadeFrom = ex && ex.played_games ? n - ex.played_games : undefined;
  const win = Math.min(10, n);
  const ma = places.map((_,i)=> i+1 < win ? null : places.slice(i+1-win, i+1).reduce((a,b)=>a+b,0)/win);
  const refPlace = ex && ex.required_place ? ex.required_place : undefined;
  drawLine(document.getElementById('trend-place'), {
    values: ma, color:'#2b74b8', invert:true, min:1, max:4, ticks:[1,2,3,4], shadeFrom,
    ref: refPlace, refLabel: refPlace ? `基準 ${refPlace}` : '',
    tickFmt: v => v.toFixed(0)+'着',
    tipFmt: (i,v) => `${i+1}半荘目時点　直近${win}半荘 平均${v.toFixed(2)}着`,
  });
}

function statsTwoPanels(p, ex){
  const c = p.career || {};
  const avoidC = c.last_avoid_rate!==null && c.last_avoid_rate!==undefined ? Math.round(c.last_avoid_rate*100)+'%' : '—';
  const avoidP = ex.last_avoid_rate!==null && ex.last_avoid_rate!==undefined ? Math.round(ex.last_avoid_rate*100)+'%' : '—';
  return `
  <div class="stats-two">
    <div class="stats-panel">
      <div class="ptitle career">累計成績</div>
      <div class="pgrid">
        <div><b>${c.games||0}</b>累計半荘数</div>
        <div><b>${fmt(c.avg_place,2)}</b>累計平均着順</div>
        <div><b>${avoidC}</b>累計ラス回避率</div>
        <div><b>${fmtPts(c.points)}</b>累計ポイント</div>
      </div>
    </div>
    <div class="stats-panel">
      <div class="ptitle period">期間内成績</div>
      <div class="pgrid">
        <div><b>${ex.played_games||0}</b>期間内半荘数</div>
        <div><b>${fmt(ex.avg_place,2)}</b>期間内平均着順</div>
        <div><b>${avoidP}</b>期間内ラス回避率</div>
        <div><b>${fmtPts(ex.total_current)}</b>期間内ポイント</div>
      </div>
    </div>
  </div>`;
}

function renderExamCard(){
  const p = players.find(x=>x.key===selectedKey);
  const el = document.getElementById('examcard');
  if(!p){
    el.innerHTML = `<div class="welcome"><div class="wt">まずは自分の名前を選びましょう</div>
      <p>上の検索欄に名前の一部を入れると候補が出ます。選ぶと、昇段に必要な成績や着順の推移が表示されます。次回からは最後に見た人が最初に出ます。</p></div>`;
    return;
  }
  const ex = p.exam;

  if(!ex || ex.status==='unranked'){
    el.innerHTML = `
      <div class="exam-head">
        <div class="exam-name">${p.name}</div>
        <div class="exam-badges"><span class="statusbadge unranked">未認定</span></div>
      </div>
      <div class="exam-msg">まだ段位が設定されていません。対局を開始すると五級からの審査がスタートします。</div>
      ${statsTwoPanels(p, {played_games:0, avg_place:null, last_avoid_rate:null, total_current:0})}
      ${trendsHtml(p)}
    `;
    renderTrends(p, null);
    return;
  }

  const tier = rankTier(ex.rank);
  const places = ex.places || [];
  const scores = ex.scores || [];
  const tiles = Array.from({length: Math.max(ex.required_games, places.length)}, (_,i)=>{
    const pl = places[i];
    if(pl) return `<div class="t p${pl}" title="${i+1}半荘目：${PLACE_LABEL[pl]}（${fmtPts(scores[i])}）">${pl}</div>`;
    return `<div class="t ${i < ex.played_games ? 'done':''}"></div>`;
  }).join('');
  const placeCount = [1,2,3,4].map(k=>places.filter(x=>x===k).length);
  
  const legend = places.length ? `<span class="placelegend">${[1,2,3,4].map(k=>
    `<span><i style="background:${PLACE_COLOR[k]}"></i>${PLACE_LABEL[k]}<b>${placeCount[k-1]}</b></span>`).join('')}</span>` : '';

  let msg = '';
  if(ex.status==='not_started'){
    msg = `直近の昇段によりこの段位の判定はリセットされました。次の対局から ${ex.required_games} 半荘の審査が始まります。`;
  } else if(ex.status==='ready'){
    msg = `規定打数 ${ex.required_games} 半荘を消化し、平均着順・平均収支ともに基準を満たしています。昇段確定ラインに到達しました。`;
  } else if(ex.status==='reset_pending'){
    msg = `規定打数 ${ex.required_games} 半荘を消化しましたが、基準未達のため次の対局から打数がリセットされ再挑戦となります。`;
  } else {
    msg = `現在 ${ex.played_games} / ${ex.required_games} 半荘を消化中です。残り ${ex.remaining_games} 半荘で基準クリアを目指しましょう。`;
  }

  const placeTag = ex.played_games>0 ? (ex.place_clear ? '<span class="mtag clear">クリア中</span>' : '<span class="mtag notclear">未達</span>') : '';
  const avgTag = ex.played_games>0 ? (ex.avg_clear ? '<span class="mtag clear">クリア中</span>' : '<span class="mtag notclear">未達</span>') : '';

  let paceHtml = '';
  if(ex.remaining_games > 0){
    if(ex.total_remaining > 0){
      paceHtml = `<div class="pacebox">残り <b>${ex.remaining_games}半荘</b> で、1半荘あたり平均 <b>${fmtPts(ex.pace_needed)}</b> 以上のペースが必要です。（不足累計: ${fmtPts(-ex.total_remaining)}）</div>`;
    } else {
      paceHtml = `<div class="pacebox">必要累計ポイントは既に確保できています。残り <b>${ex.remaining_games}半荘</b> を消化すれば昇段判定です（平均着順の基準維持にご注意ください）。</div>`;
    }
    paceHtml += placeExampleHtml(ex);
  }

  el.innerHTML = `
    <div class="exam-head">
      <div class="exam-name">${p.name}</div>
      <div class="exam-badges">
        <span class="rankbadge" style="background:${tier}">${ex.rank} 昇段審査中</span>
        <span class="statusbadge ${ex.status}">${STATUS_LABEL[ex.status]}</span>
      </div>
    </div>
    <div class="exam-msg">${msg}</div>
    <div class="tilewall">
      <div class="lbl"><span>半荘進捗　${ex.played_games} / ${ex.required_games}</span>${legend}</div>
      ${tiles}
    </div>
    <div class="metrics">
      <div class="metric">
        <div class="mlabel"><span>平均着順（${ex.required_place}以内）</span>${placeTag}</div>
        <div class="mval">${fmt(ex.avg_place,3)}<span class="req">/ 基準 ${ex.required_place}</span></div>
      </div>
      <div class="metric">
        <div class="mlabel"><span>平均収支（${ex.required_avg}以上）</span>${avgTag}</div>
        <div class="mval">${fmt(ex.avg_score,1)}<span class="req">/ 基準 ${ex.required_avg}</span></div>
      </div>
    </div>
    ${paceHtml}
    ${simulatorHtml(ex)}
    ${statsTwoPanels(p, ex)}
    ${trendsHtml(p)}
  `;
  bindSimulator(ex, p);
  renderTrends(p, ex);
}

const sortOptions = [
  {key:'rank_value', label:'段位順', dir:-1},
  {key:'points', label:'ポイント順', dir:-1},
  {key:'avg_place', label:'平均着順順', dir:1},
  {key:'games', label:'半荘数順', dir:-1},
];
let sortKey = 'rank_value';
let sortDir = -1;

function renderSortButtons(){
  const el = document.getElementById('sortbtns');
  el.innerHTML = sortOptions.map(o=>`<button data-key="${o.key}" data-dir="${o.dir}">${o.label}</button>`).join('');
  el.querySelectorAll('button').forEach(btn=>{
    if(btn.dataset.key===sortKey) btn.classList.add('active');
    btn.addEventListener('click', ()=>{
      sortKey = btn.dataset.key; sortDir = parseInt(btn.dataset.dir);
      el.querySelectorAll('button').forEach(b=>b.classList.remove('active'));
      btn.classList.add('active');
      renderRows();
    });
  });
}

function highlightRow(){
  document.querySelectorAll('.row[data-key]').forEach(r=>{
    r.classList.toggle('selected', r.dataset.key===selectedKey);
  });
}

function playerSortValue(p, key){
  if(key==='rank_value') return p.rank_value;
  if(key==='points') return p.career.points;
  if(key==='avg_place') return p.career.avg_place;
  if(key==='games') return p.career.games;
  return 0;
}

// ---- 審査状況フィルタ ----
let statusFilter = 'all';
const FILTERS = [
  {key:'all', label:'すべて'}, {key:'ready', label:'合格ライン到達'}, {key:'on_track', label:'基準クリア中'},
  {key:'behind', label:'追い上げ中'}, {key:'reset_pending', label:'要再挑戦'}, {key:'not_started', label:'審査開始前'},
];
function renderFilters(){
  const el = document.getElementById('filterbtns');
  const count = k => k==='all' ? activePlayers.length : activePlayers.filter(p=>(p.exam?p.exam.status:'unranked')===k).length;
  el.innerHTML = FILTERS.filter(f=>f.key==='all'||count(f.key)>0)
    .map(f=>`<button type="button" data-key="${f.key}" class="${f.key===statusFilter?'active':''}">${f.label}<b>${count(f.key)}</b></button>`).join('');
  el.querySelectorAll('button').forEach(b=>b.addEventListener('click', ()=>{
    statusFilter = b.dataset.key;
    el.querySelectorAll('button').forEach(x=>x.classList.toggle('active', x===b));
    renderRows();
  }));
}

function renderRows(){
  const q = document.getElementById('search').value.trim().toLowerCase();
  let list = activePlayers.filter(p=>p.name.toLowerCase().includes(q))
    .filter(p=> statusFilter==='all' || (p.exam ? p.exam.status : 'unranked')===statusFilter);
  list.sort((a,b)=>{
    let av = playerSortValue(a, sortKey), bv = playerSortValue(b, sortKey);
    if(av===null) av = -Infinity;
    if(bv===null) bv = -Infinity;
    if(av===bv) return (b.career.points||0) - (a.career.points||0);
    return (av-bv) * sortDir;
  });
  const rowsEl = document.getElementById('rows');
  if(list.length===0){ rowsEl.innerHTML = '<div class="empty">該当する打ち手が見つかりません</div>'; return; }
  rowsEl.innerHTML = list.map((p,i)=>{
    const tier = rankTier(p.rank);
    const status = p.exam ? p.exam.status : 'unranked';
    return `
    <div class="row" data-key="${p.key}">
      <div class="rk">${i+1}</div>
      <div class="nm">${p.name}</div>
      <div><span class="badge" style="background:${tier}">${p.rank}</span></div>
      <div class="num col-place"><b>${fmt(p.career.avg_place,2)}</b></div>
      <div class="num col-games"><b>${p.career.games||0}</b></div>
      <div class="st col-status ${status}">${STATUS_LABEL[status]}</div>
    </div>`;
  }).join('');
  rowsEl.querySelectorAll('.row').forEach(r=>{
    r.addEventListener('click', ()=> selectPlayer(r.dataset.key, {scroll:true}));
  });
  highlightRow();
}

document.getElementById('search').addEventListener('input', renderRows);

renderStatbar();
renderPicker();
renderNews();
renderExamCard();
renderFilters();
renderSortButtons();
renderRows();
if(new URLSearchParams(location.search).get('p')){
  setTimeout(()=>document.getElementById('examcard').scrollIntoView({block:'start'}), 50);
}
let _rz; window.addEventListener('resize', ()=>{ clearTimeout(_rz); _rz=setTimeout(()=>{
  const p = players.find(x=>x.key===selectedKey); if(p) renderTrends(p, p.exam && p.exam.status!=='unranked' ? p.exam : null);
}, 150); });
</script>
</body>
</html>
"""


def build_html(data):
    data_json = json.dumps(data, ensure_ascii=False)
    generated_at = datetime.datetime.now().strftime("%Y年%m月%d日 %H:%M 更新")
    date_only = datetime.datetime.now().strftime("%Y年%m月%d日時点")
    return (TEMPLATE
            .replace("__DATA_JSON__", data_json)
            .replace("__GENERATED_AT__", generated_at)
            .replace("__DATE_ONLY__", date_only))


def main():
    args = sys.argv[1:]
    prev_path = None
    if "--prev" in args:
        i = args.index("--prev")
        prev_path = args[i + 1]
        del args[i:i + 2]
    if len(args) < 2:
        print("使い方: python3 generate_dashboard.py <入力xlsx> <出力html> [--prev <前回のindex.html>]")
        print("  --prev 省略時は、出力先に既にあるhtmlを前回分として比較に使う")
        sys.exit(1)
    xlsx_path, out_path = args[0], args[1]
    data, warnings = extract(xlsx_path)
    prev = load_prev_data(prev_path or out_path)
    data["changes"] = compute_changes(data, prev)
    html = build_html(data)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"生成完了: {out_path}（{len(data['players'])}名分）")
    ch = data["changes"]
    print(f"[今回の更新] 比較元: {ch.get('since') or 'なし'} / 昇段 {len(ch['promoted'])}名・"
          f"新メンバー {len(ch['new'])}名・合格ライン到達 {len(ch['ready'])}名・昇段リーチ {len(ch['reach'])}名")
    if warnings:
        print("--- 確認事項 ---")
        for w in warnings:
            print("・" + w)


if __name__ == "__main__":
    main()
