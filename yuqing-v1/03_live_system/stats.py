# -*- coding: utf-8 -*-
"""统计引擎：历史基线（第一阶段 data.js） + 实时增量（incidents）合并"""
import copy
import json
import re
from datetime import datetime

import db
from config import BASE_DATA_JS


def parse_data_js(path):
    with open(path, encoding="utf-8") as f:
        s = f.read()
    s = re.sub(r"^.*?window\.DASH_DATA\s*=\s*", "", s, count=1, flags=re.S)
    s = s.strip().rstrip(";").strip()
    return json.loads(s)


def load_base_data():
    data = db.get_meta("base_data")
    if data is None:
        data = parse_data_js(BASE_DATA_JS)
        db.set_meta("base_data", data)
    return data


def now_iso():
    return datetime.now().strftime("%Y-%m-%dT%H:%M:%S")


def _add_num(d, key, n):
    d[key] = d.get(key, 0) + n


def compute_live_stats():
    rows = db.query_all(
        "SELECT * FROM incidents WHERE status='accepted' AND origin<>'history' ORDER BY id DESC"
    )
    acc = {
        "total": 0, "support": 0, "neutral": 0, "suggest": 0,
        "non_support": 0, "other": 0, "region_total": 0, "region_support": 0,
        "minority": 0,
    }
    platforms = {}
    regions = {}
    provinces = {}
    langs = {}
    trend = {}
    nonsup = {}
    hot = []
    for r in rows:
        acc["total"] += 1
        bucket = r["attitude_bucket"] or "other"
        _add_num(acc, bucket, 1)
        if r["region"]:
            acc["region_total"] += 1
            if bucket == "support":
                acc["region_support"] += 1
        if r["is_minority"]:
            acc["minority"] += 1

        g = r["platform_group"] or "其他"
        p = platforms.setdefault(g, {
            "name": g, "total": 0, "support": 0, "neutral": 0, "qa": 0, "worry": 0,
            "criticism": 0, "complaint": 0, "implement": 0, "fairness": 0,
            "discrimination": 0, "suggest": 0, "unknownLaw": 0, "other": 0,
            "supportRate": 0.0, "nonSupport": 0,
        })
        p["total"] += 1
        p["support"] += 1 if bucket == "support" else 0
        p["neutral"] += 1 if bucket == "neutral" else 0
        p["suggest"] += 1 if bucket == "suggest" else 0
        if bucket == "non_support":
            p["nonSupport"] += 1
            cat = r["issue_category"]
            if cat in ("咨询疑问", "担忧影响", "明确批评", "投诉维权", "实施问题", "公平争议", "歧视偏见", "不了解该法律"):
                p[{"咨询疑问": "qa", "担忧影响": "worry", "明确批评": "criticism", "投诉维权": "complaint",
                   "实施问题": "implement", "公平争议": "fairness", "歧视偏见": "discrimination",
                   "不了解该法律": "unknownLaw"}[cat]] += 1
            else:
                p["other"] += 1
            _add_num(nonsup, cat or "其他/未分类", 1)

        if r["province"] or r["region"]:
            name = r["province"] or r["region"] or "未分组"
            reg = regions.setdefault(name, {
                "name": name, "total": 0, "support": 0, "neutral": 0, "qa": 0, "worry": 0,
                "criticism": 0, "complaint": 0, "implement": 0, "fairness": 0,
                "discrimination": 0, "pending": 0, "provinces": [name],
                "sourceGroup": r["region"] or "",
            })
            reg["total"] += 1
            reg["support"] += 1 if bucket == "support" else 0
            reg["neutral"] += 1 if bucket == "neutral" else 0
            cat = r["issue_category"]
            if cat in ("咨询疑问", "担忧影响", "明确批评", "投诉维权", "实施问题", "公平争议", "歧视偏见"):
                reg[{"咨询疑问": "qa", "担忧影响": "worry", "明确批评": "criticism", "投诉维权": "complaint",
                     "实施问题": "implement", "公平争议": "fairness", "歧视偏见": "discrimination"}[cat]] += 1

        if r["province"]:
            name = r["province"]
            prov = provinces.setdefault(name, {
                "name": name, "short": name, "value": 0, "total": 0, "support": 0,
                "neutral": 0, "qa": 0, "worry": 0, "criticism": 0, "complaint": 0,
                "implement": 0, "fairness": 0, "discrimination": 0, "suggest": 0, "other": 0,
            })
            prov["value"] += 1
            prov["total"] += 1
            prov["support"] += 1 if bucket == "support" else 0
            prov["neutral"] += 1 if bucket == "neutral" else 0
            cat = r["issue_category"]
            if cat in ("咨询疑问", "担忧影响", "明确批评", "投诉维权", "实施问题", "公平争议", "歧视偏见", "不了解该法律"):
                prov[{"咨询疑问": "qa", "担忧影响": "worry", "明确批评": "criticism", "投诉维权": "complaint",
                      "实施问题": "implement", "公平争议": "fairness", "歧视偏见": "discrimination",
                      "不了解该法律": "unknownLaw"}[cat]] += 1
            elif bucket == "suggest":
                prov["suggest"] += 1
            else:
                prov["other"] += 1

        lang = r["language"] or "中文"
        lv = langs.get(lang, {"name": lang, "value": 0})
        lv["value"] += 1
        langs[lang] = lv

        d = (r["published_at"] or "")[:10]
        if d:
            trend[d] = trend.get(d, 0) + 1

        if r["likes"]:
            hot.append({
                "platform": r["platform"] or r["platform_group"] or "",
                "account": r["account"] or "",
                "title": (r["text"] or "")[:60],
                "likes": r["likes"],
                "comments": r["comments"] or 0,
                "shares": r["shares"] or 0,
                "date": d,
            })

    hot.sort(key=lambda x: x["likes"], reverse=True)
    pending = db.query_one("SELECT COUNT(*) AS n FROM incidents WHERE status='pending'")["n"]
    return {
        "counts": acc,
        "platforms": list(platforms.values()),
        "regions": list(regions.values()),
        "provinces": list(provinces.values()),
        "languages": list(langs.values()),
        "trend": trend,
        "nonSupport": nonsup,
        "hot": hot,
        "pending": pending,
        "rows": rows,
    }


def row_to_quote(r):
    return {
        "platform": r["platform"] or r["platform_group"] or "",
        "region": r["region"] or r["province"] or "",
        "group": r["region"] or "",
        "attitude": r["attitude"] or "",
        "text": r["text"] or "",
        "date": (r["published_at"] or "")[:10],
        "source": r["source"] or r["origin"] or "",
        "src": r["source"] or r["origin"] or "",
        "account": r["account"] or "",
        "language": r["language"] or "中文",
        "collectedAt": r["collected_at"] or "",
        "isLive": True,
        "id": r["id"],
    }


def merge_into_list(base_list, extra_list, key, sum_keys):
    out = copy.deepcopy(base_list)
    by = {}
    for item in out:
        by[item.get(key, "")] = item
    for extra in extra_list:
        name = extra.get(key, "")
        if not name:
            continue
        if name in by:
            for k in sum_keys:
                by[name][k] = (by[name].get(k, 0) or 0) + (extra.get(k, 0) or 0)
        else:
            item = copy.deepcopy(extra)
            item.setdefault("total", 0)
            by[name] = item
    return list(by.values())


def build_bootstrap():
    base = copy.deepcopy(load_base_data())
    live = compute_live_stats()
    c = live["counts"]

    # 顶部 KPI
    ts = base.setdefault("topStats", {})
    ts["totalOpinions"] = (ts.get("totalOpinions") or 0) + c["total"]
    ts["supportCount"] = (ts.get("supportCount") or 0) + c["support"]
    ts["regionOpinions"] = (ts.get("regionOpinions") or 0) + c["region_total"]
    ts["regionSupport"] = (ts.get("regionSupport") or 0) + c["region_support"]
    ts["nonSupport"] = (ts.get("nonSupport") or 0) + c["non_support"]
    ts["minorityLang"] = (ts.get("minorityLang") or 0) + c["minority"]
    if ts.get("totalOpinions"):
        ts["supportRate"] = round((ts.get("supportCount") or 0) / ts["totalOpinions"] * 100, 2)
    if ts.get("regionOpinions"):
        ts["regionSupportRate"] = round((ts.get("regionSupport") or 0) / ts["regionOpinions"] * 100, 2)

    # 平台 / 地区 / 省级 / 语言
    base["platforms"] = merge_into_list(
        base.get("platforms", []), live["platforms"], "name",
        ["total", "support", "neutral", "qa", "worry", "criticism", "complaint",
         "implement", "fairness", "discrimination", "suggest", "unknownLaw", "other", "nonSupport"],
    )
    base["regions"] = merge_into_list(
        base.get("regions", []), live["regions"], "name",
        ["total", "support", "neutral", "qa", "worry", "criticism", "complaint",
         "implement", "fairness", "discrimination", "pending"],
    )
    base["provinces"] = merge_into_list(
        base.get("provinces", []), live["provinces"], "name",
        ["value", "total", "support", "neutral", "qa", "worry", "criticism", "complaint",
         "implement", "fairness", "discrimination", "suggest", "other"],
    )
    base["languagePlatform"] = merge_into_list(
        base.get("languagePlatform", []), live["languages"], "name", ["value"],
    )

    # 态度构成
    att_extra = {
        "支持认可": c["support"], "中性信息": c["neutral"], "参与建议": c["suggest"],
        "非支持/非肯定": c["non_support"],
    }
    for item in base.setdefault("attitude", {}).get("macro", []):
        item["value"] = (item.get("value") or 0) + att_extra.get(item["name"], 0)
    detail_extra = {
        "支持认可": c["support"], "中性信息": c["neutral"], "参与建议": c["suggest"],
        "咨询疑问": live["nonSupport"].get("咨询疑问", 0), "担忧影响": live["nonSupport"].get("担忧影响", 0),
        "明确批评": live["nonSupport"].get("明确批评", 0), "投诉维权": live["nonSupport"].get("投诉维权", 0),
        "实施问题": live["nonSupport"].get("实施问题", 0), "公平争议": live["nonSupport"].get("公平争议", 0),
        "歧视偏见": live["nonSupport"].get("歧视偏见", 0), "不了解该法律": live["nonSupport"].get("不了解该法律", 0),
    }
    for item in base["attitude"].setdefault("detail", []):
        item["value"] = (item.get("value") or 0) + detail_extra.get(item["name"], 0)

    # 非支持构成
    base["nonSupport"] = merge_into_list(
        base.get("nonSupport", []),
        [{"name": k, "value": v} for k, v in live["nonSupport"].items()],
        "name", ["value"],
    )

    # 时间趋势
    trend_by = {}
    for t in base.get("trend", []):
        trend_by[t["date"]] = trend_by.get(t["date"], 0) + (t.get("value") or 0)
    for d, v in live["trend"].items():
        trend_by[d] = trend_by.get(d, 0) + v
    base["trend"] = [{"date": d, "value": v} for d, v in sorted(trend_by.items())]

    # 原话池（历史 + 实时）
    base["quotes"] = (base.get("quotes") or []) + [row_to_quote(r) for r in live["rows"]]

    # 高热内容
    hot = (base.get("hotTop") or []) + live["hot"]
    seen = set()
    deduped = []
    for h in hot:
        key = (h.get("account", ""), h.get("title", ""), h.get("likes", 0))
        if key in seen:
            continue
        seen.add(key)
        deduped.append(h)
    deduped.sort(key=lambda x: x.get("likes", 0), reverse=True)
    base["hotTop"] = deduped[:9]

    base["generatedAt"] = now_iso()
    base["live"] = {
        "serverTime": now_iso(),
        "pending": live["pending"],
        "incidents": [row_to_quote(r) for r in live["rows"][:60]],
        "counts": {
            "total": c["total"], "support": c["support"], "nonSupport": c["non_support"],
            "region": c["region_total"], "minorityLang": c["minority"],
        },
    }
    return base
