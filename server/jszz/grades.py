"""成绩分析：真实成绩表 → 统计指标（供 grades 报告 docx / MCP）。

列检测适配 U 盘真实成绩表三种形态：
  1. 多考试长表  `考试|班级|姓名|语文|数学|…`（1870 行多次考试堆叠）
  2. 调考成绩单  `班级|姓名|学籍号|考号|总分转换分|市名次|各科+名次+转换分+等级(B+)`
  3. 单次考试    `序号|姓名|九科+日语俄语|总分` / `班级|姓名|期中成绩|周六成绩|总和`

指标（照教师质量分析真实关注点）：均分/最高/最低/中位/标准差/及格率/优秀率/
分数段分布/名次榜/单科榜/多次考试趋势(进步退步)/等级分布。
"""
from __future__ import annotations

import json
import math
import re
from pathlib import Path
from typing import Any

SUBJECTS = ["语文", "数学", "英语", "物理", "化学", "生物", "政治", "历史", "地理",
            "日语", "俄语", "总分", "总和"]
_NAME_COLS = {"姓名", "名字", "学生", "考生"}
_CLASS_COLS = {"班级", "班", "行政班"}
_EXAM_COLS = {"考试", "考次", "学期", "场次"}
_TOTAL_COLS = {"总分", "总和", "总成绩", "合计", "总分转换分"}
# ID/序号列绝不当分数（学籍号/考号是长整数，被求均分会出 2.02e18 这种荒谬值）
_ID_COLS = {"学籍号", "考号", "考籍号", "序号", "编号", "学号", "id", "ID", "座位号", "考场"}
_RANK_HINT = ("名次", "排名", "位次")
_GRADE_HINT = ("等级", "等第", "档次")


# ---------- 载入 ----------
def load_scores(source: Any) -> dict[str, Any]:
    """xlsx 路径 / JSON 字符串 / 已解析的行 dict 列表 → 标准化 {rows, columns}。

    返回 columns: {name, class_, exam, total, subjects:[…], ranks:[…], grades:[…]}。
    """
    if isinstance(source, (list, tuple)):
        rows = [dict(r) for r in source]
    elif isinstance(source, (str, Path)) and str(source).lstrip().startswith(("[", "{")):
        rows = json.loads(str(source))
    elif isinstance(source, (str, Path)):
        rows = _read_xlsx(str(source))
    else:
        rows = [dict(r) for r in source]
    rows = [{str(k).strip(): v for k, v in r.items()} for r in rows]
    if not rows:
        return {"ok": False, "error": "没有数据行", "rows": [], "columns": {}}
    cols = _detect_columns(rows)
    return {"ok": True, "rows": rows, "columns": cols,
            "n_rows": len(rows), "fields": sorted(rows[0].keys())}


def _read_xlsx(path: str) -> list[dict]:
    try:
        import openpyxl
    except Exception:  # noqa: BLE001
        return []
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.worksheets[0]
    raw = [[c for c in r] for r in ws.iter_rows(values_only=True)]
    # 找表头行：第一行里含「姓名」或「班级」或任意科目名的行
    header_i = 0
    for i, r in enumerate(raw[:5]):
        cells = [str(c).strip() if c is not None else "" for c in r]
        if any(c in _NAME_COLS or c in _CLASS_COLS or c in SUBJECTS for c in cells):
            header_i = i
            break
    header = [str(c).strip() if c is not None else f"col{j}"
              for j, c in enumerate(raw[header_i])]
    out = []
    for r in raw[header_i + 1:]:
        if all(c is None or str(c).strip() == "" for c in r):
            continue
        row = {}
        for j, val in enumerate(r):
            key = header[j] if j < len(header) else f"col{j}"
            if key in ("col0", "") or key.startswith("col") and not str(val or "").strip():
                continue
            row[key] = val
        if row:
            out.append(row)
    return out


def _detect_columns(rows: list[dict]) -> dict:
    keys = list(rows[0].keys())
    def pick(names):
        for k in keys:
            if k in names:
                return k
        return ""
    name = pick(_NAME_COLS)
    class_ = pick(_CLASS_COLS)
    exam = pick(_EXAM_COLS)
    total = pick(_TOTAL_COLS)
    subjects, ranks, grades = [], [], []
    for k in keys:
        if k in (name, class_, exam, total):
            continue
        if k in _ID_COLS or any(h in k for h in _ID_COLS):
            continue  # 学籍号/考号/序号等 ID 列不当分数
        if any(h in k for h in _RANK_HINT):
            ranks.append(k)
        elif any(h in k for h in _GRADE_HINT):
            grades.append(k)
        elif k in SUBJECTS or re.search(r"(成绩|分数)$", k) or _looks_score(rows, k):
            subjects.append(k)
    return {"name": name or keys[0], "class_": class_, "exam": exam,
            "total": total or (subjects[-1] if subjects else ""),
            "subjects": subjects, "ranks": ranks, "grades": grades}


def _looks_score(rows: list[dict], key: str) -> bool:
    """非固定科目名的数值列（如「期中成绩」「周六成绩」）也当分数列。"""
    vals = [r.get(key) for r in rows[:30]]
    nums = [v for v in vals if _num(v) is not None]
    return len(nums) >= max(3, len(vals) // 3)


def _num(v) -> float | None:
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace("／", "/")
    if not s or s in ("·", "-", "—", "缺考", "null", "None"):
        return None
    try:
        return float(s)
    except ValueError:
        return None


# ---------- 统计 ----------
def _full_mark(vals: list[float]) -> int:
    mx = max(vals) if vals else 0
    return 150 if mx > 100 else 100


def _stats(vals: list[float], full: int | None = None) -> dict:
    if not vals:
        return {"n": 0}
    full = full or _full_mark(vals)
    vals = sorted(vals)
    n = len(vals)
    mean = sum(vals) / n
    median = (vals[n // 2] if n % 2 else (vals[n // 2 - 1] + vals[n // 2]) / 2)
    std = math.sqrt(sum((v - mean) ** 2 for v in vals) / n)
    pa, ex = 0.6 * full, 0.85 * full
    bands = {"≥%g" % round(0.9 * full, 1): 0, "%g~%g" % (round(0.8 * full, 1), round(0.9 * full, 1)): 0,
             "%g~%g" % (round(0.7 * full, 1), round(0.8 * full, 1)): 0,
             "%g~%g" % (round(0.6 * full, 1), round(0.7 * full, 1)): 0,
             "<%g" % round(0.6 * full, 1): 0}
    for v in vals:
        if v >= 0.9 * full:
            bands["≥%g" % round(0.9 * full, 1)] += 1
        elif v >= 0.8 * full:
            bands["%g~%g" % (round(0.8 * full, 1), round(0.9 * full, 1))] += 1
        elif v >= 0.7 * full:
            bands["%g~%g" % (round(0.7 * full, 1), round(0.8 * full, 1))] += 1
        elif v >= 0.6 * full:
            bands["%g~%g" % (round(0.6 * full, 1), round(0.7 * full, 1))] += 1
        else:
            bands["<%g" % round(0.6 * full, 1)] += 1
    return {
        "n": n, "full": full, "mean": round(mean, 2), "median": round(median, 2),
        "max": round(vals[-1], 2), "min": round(vals[0], 2), "std": round(std, 2),
        "pass_rate": round(100 * sum(1 for v in vals if v >= pa) / n, 1),
        "excellence_rate": round(100 * sum(1 for v in vals if v >= ex) / n, 1),
        "pass_line": pa, "excellence_line": ex,
        "bands": bands,
    }


def analyze(rows: list[dict], columns: dict | None = None,
            top_n: int = 10, full_marks: dict | None = None) -> dict:
    """全量分析：各科统计 + 分数段 + 名次榜 + 多次考试趋势 + 等级分布。"""
    if columns is None:
        columns = _detect_columns(rows)
    full_marks = full_marks or {}
    nm, cl, exm = columns["name"], columns["class_"], columns["exam"]
    subjects = columns["subjects"]
    total_key = columns["total"]
    per_subj: dict[str, dict] = {}
    for s in subjects:
        vals = [v for v in (_num(r.get(s)) for r in rows) if v is not None]
        per_subj[s] = _stats(vals, full_marks.get(s))
    total_vals = [v for v in (_num(r.get(total_key)) for r in rows) if v is not None] \
        if total_key else []
    total_stats = _stats(total_vals, full_marks.get(total_key)) if total_vals else {"n": 0}

    # 名次榜（按总分，无总分用第一科）；多考次长表按「每人最近一次」去重，
    # 否则同一学生跨 N 个考次占 N 个名次（崔雪×4），榜失去意义
    rank_key = total_key if total_vals else (subjects[0] if subjects else "")
    latest_rows = rows
    if exm:
        order: dict[str, int] = {}
        for r in rows:
            e = str(r.get(exm, ""))
            if e and e not in order:
                order[e] = len(order)
        by_name: dict[str, dict] = {}
        for r in rows:
            nm_v = str(r.get(nm, ""))
            e = str(r.get(exm, ""))
            cur = by_name.get(nm_v)
            if cur is None or order.get(e, -1) >= order.get(str(cur.get(exm, "")), -1):
                by_name[nm_v] = r
        latest_rows = list(by_name.values())
    scored = [(str(r.get(nm, "")), str(r.get(cl, "")), _num(r.get(rank_key)))
              for r in latest_rows]
    scored = [t for t in scored if t[2] is not None]
    scored.sort(key=lambda t: t[2], reverse=True)
    top = [{"name": t[0], "class_": t[1], "score": t[2]} for t in scored[:top_n]]
    bottom = [{"name": t[0], "class_": t[1], "score": t[2]} for t in scored[-top_n:][::-1]]

    # 单科榜首（同样按每人最近一次去重）
    subject_tops = {}
    for s in subjects:
        sv = [(str(r.get(nm, "")), _num(r.get(s))) for r in latest_rows]
        sv = [t for t in sv if t[1] is not None]
        sv.sort(key=lambda t: t[1], reverse=True)
        subject_tops[s] = [{"name": t[0], "score": t[1]} for t in sv[:3]]

    # 班级对比
    class_stats = {}
    if cl:
        groups: dict[str, list[float]] = {}
        for r in rows:
            v = _num(r.get(rank_key))
            if v is not None:
                groups.setdefault(str(r.get(cl, "")), []).append(v)
        for cname, vals in sorted(groups.items()):
            class_stats[cname] = _stats(vals, full_marks.get(rank_key))

    # 多次考试趋势（进步/退步）
    trend = None
    if exm:
        by_student: dict[str, list[tuple[str, float]]] = {}
        for r in rows:
            v = _num(r.get(rank_key))
            if v is None:
                continue
            by_student.setdefault(str(r.get(nm, "")), []).append((str(r.get(exm, "")), v))
        # 考次顺序：按首次出现顺序
        exam_order: list[str] = []
        for r in rows:
            e = str(r.get(exm, ""))
            if e and e not in exam_order:
                exam_order.append(e)
        deltas = []
        for name, series in by_student.items():
            if len(series) < 2:
                continue
            order = {e: i for i, e in enumerate(exam_order)}
            series.sort(key=lambda t: order.get(t[0], 0))
            deltas.append({"name": name, "first": series[0][1], "last": series[-1][1],
                           "delta": round(series[-1][1] - series[0][1], 2),
                           "n": len(series)})
        deltas.sort(key=lambda d: d["delta"], reverse=True)
        exam_means = {}
        for e in exam_order:
            vals = [_num(r.get(rank_key)) for r in rows if str(r.get(exm, "")) == e]
            vals = [v for v in vals if v is not None]
            if vals:
                exam_means[e] = round(sum(vals) / len(vals), 2)
        trend = {"exams": exam_order, "exam_means": exam_means,
                 "improved": deltas[:5], "declined": deltas[-5:][::-1],
                 "n_students": len(deltas)}

    # 等级分布
    grade_dist = {}
    for g in columns["grades"]:
        cnt: dict[str, int] = {}
        for r in rows:
            key = str(r.get(g, "")).strip()
            if key and key not in ("·", "None", ""):
                cnt[key] = cnt.get(key, 0) + 1
        if cnt:
            grade_dist[g] = dict(sorted(cnt.items()))

    # 弱科识别 + 提升推荐（按均分/及格率/优秀率综合判定，照写作工艺短句采分）
    weak = _weak_subjects(per_subj, total_stats)
    recommendations = _recommendations(per_subj, weak, trend)

    return {
        "ok": True, "n_rows": len(rows), "columns": columns,
        "subjects": per_subj, "total": total_stats, "top": top, "bottom": bottom,
        "subject_tops": subject_tops, "class_stats": class_stats,
        "trend": trend, "grade_dist": grade_dist, "rank_key": rank_key,
        "weak_subjects": weak, "recommendations": recommendations,
    }


def _weak_subjects(per_subj: dict, total_stats: dict) -> list[dict]:
    """弱科判定：均分率（均分/满分）低于阈值，或及格率/优秀率明显偏低。

    返回按「弱度」排序的 [{subject, mean, full, mean_rate, pass_rate, excellence_rate, reason}]。
    """
    out = []
    for s, st in per_subj.items():
        if not st.get("n") or not st.get("full"):
            continue
        # 总分列不当弱科（它是合计）
        if s in ("总分", "总和", "总成绩", "合计"):
            continue
        rate = st.get("mean", 0) / st["full"] if st.get("full") else 0
        pr = st.get("pass_rate", 100)
        er = st.get("excellence_rate", 100)
        # 弱度：均分率权重高，及格率/优秀率作修正
        weakness = (1 - rate) * 0.6 + (1 - pr / 100) * 0.25 + (1 - er / 100) * 0.15
        reasons = []
        if rate < 0.6:
            reasons.append(f"均分率仅 {rate * 100:.0f}%（{st['mean']}/{st['full']}）")
        if pr < 60:
            reasons.append(f"及格率 {pr}% 偏低")
        if er < 15:
            reasons.append(f"优秀率 {er}% 偏低")
        if reasons:
            out.append({"subject": s, "mean": st.get("mean"), "full": st.get("full"),
                        "mean_rate": round(rate * 100, 1), "pass_rate": pr,
                        "excellence_rate": er, "weakness": round(weakness, 3),
                        "reason": "，".join(reasons)})
    out.sort(key=lambda x: x["weakness"], reverse=True)
    return out


def _recommendations(per_subj: dict, weak: list[dict], trend: dict | None) -> list[str]:
    """提升推荐：针对弱科给可操作措施（照写作工艺：短句、带采分点思维、不口语）。

    每条带各科**具体短板数字**（均分率/及格率/优秀率里最刺眼的那个），
    不让多科落进同一模板句读起来像机器套话。
    """
    recs = []
    for i, w in enumerate(weak[:3]):
        s = w["subject"]
        rank = ["首要", "其次", "第三"][i] if i < 3 else "另外"
        # 挑该科最弱短板做切入点；但优秀率常并列 0（多科都是 0），此时改用均分率
        # （29/34/47 各不同），否则三科都指向"优秀率0"→ 动作雷同读着像机器套话。
        gaps = [(w["mean_rate"], "均分率"), (w["pass_rate"], "及格率"),
                (w["excellence_rate"], "优秀率")]
        gaps.sort(key=lambda x: x[0])
        worst_val, worst_name = gaps[0]
        if worst_name == "优秀率" and w["excellence_rate"] == 0:
            worst_val, worst_name = w["mean_rate"], "均分率"
        # 动作文案 + 结尾短语都随序号轮换，杜绝三科同款（含结尾一字不差）
        fills = ["回归教材抓基本概念与常考题型，每日 10 分钟基础默写",
                 "分层布置作业，基础题全员过关、中档题每日 1 题",
                 "梳理核心公式与定理，配合错题重做补齐断层"]
        tails = ["一月内冲及格率 60%", "先把及格线稳住再说", "目标及格率过 6 成"]
        if w["pass_rate"] < 40:
            recs.append(f"{rank}突破 {s}：{worst_name}仅 {worst_val}%"
                        f"（{w['mean']}/{w['full']}），{fills[i % 3]}，{tails[i % 3]}")
        elif w["pass_rate"] < 60:
            recs.append(f"{rank}突破 {s}：及格率 {w['pass_rate']}% 未过半，"
                        f"{fills[(i + 1) % 3]}，{tails[(i + 1) % 3]}")
        elif w["excellence_rate"] < 15:
            recs.append(f"{rank}提优 {s}：优秀率仅 {w['excellence_rate']}%，"
                        f"强化综合与压轴题专练，每周 2 道精练冲 15%")
        else:
            recs.append(f"{rank}巩固 {s}：针对 {worst_name} 薄弱知识点专项突破，"
                        f"建错题本定期重做")
    # 趋势建议
    if trend and trend.get("declined"):
        n = len(trend["declined"])
        recs.append(f"退步学生跟踪：{n} 名学生近期退步，个别谈话 + 错题重做，一周内复测")
    # 通识建议（写作工艺：规范答题）
    recs.append("规范答题：按采分点分行书写，杜绝口语化表达，减少非知识性失分")
    return recs
