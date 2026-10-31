#!/usr/bin/env python3
"""写作工艺挖掘：从全量课件/学案 txt 里挖「怎么写」的语言标记。
输出：analysis/full4/writecraft.json + writecraft_digest.md
维度：导入/设问/过渡/归纳口诀/警示易错/互动指令/讲评/学科气质
"""
import json, re, glob, os, collections

TXT = "analysis/full4/txt"

# 学科推断（按文件名）
SUBJ = {
    "语文": ["语文", "文言", "古诗", "作文", "阅读", "默写", "《", "师说", "劝学", "赤壁", "烛之武", "种树", "石钟山", "诗歌", "小说", "散文"],
    "数学": ["数学", "函数", "导数", "数列", "三角", "向量", "概率", "立体几何", "解析几何", "不等式", "集合", "导学案", "圆锥"],
    "英语": ["英语", "Unit", "U1", "U2", "U3", "U4", "U5", "U6", "B1U", "B2U", "B3U", "阅读", "完形", "写作", "语法", "It句型", "时态", "从句", "作文", "晨诵", "晚写"],
    "物理": ["物理", "动量", "受力", "电场", "磁场", "力学", "运动", "能量", "振动", "波", "光学", "原子", "电路", "实验"],
    "化学": ["化学", "反应", "离子", "有机", "氧化", "还原", "催化", "电极", "电池", "速率", "平衡", "元素", "金属", "非金属", "溶液", "电解"],
    "生物": ["生物", "细胞", "基因", "遗传", "光合", "呼吸", "生态", "进化", "神经", "激素", "酶", "DNA", "染色体"],
    "政治": ["政治", "经济", "哲学", "文化", "法治", "时政", "唯物", "辩证", "价值", "制度"],
    "历史": ["历史", "朝代", "近代", "古代", "史", "改革", "革命", "世界史", "中国史", "纲要"],
    "地理": ["地理", "气候", "地形", "大气", "水文", "洋流", "地质", "人口", "城市", "农业", "工业", "区域", "等高线", "热力环流"],
}
def guess_subj(rel):
    for s, kws in SUBJ.items():
        for k in kws:
            if k.lower() in rel.lower():
                return s
    return "未分类"

# 写作语言维度 → 正则
CATS = {
    "导入开场": r"(同学们|同学|上节课|上节|这节课|今天我们|本节课|我们(来|一起)|先(回顾|看|想)|请(看|想|回忆)|导入|新课|开门见山|情境|想象一下|设想)",
    "设问引导": r"(什么是|为什么(呢|呀)?|怎么(样|做)?|如何|怎样|是不是|对不对|有没有|能不能|你会(吗|么)?|你能(吗|么)?|试问|请思考|想一想|思考一下|讨论一下|试一试|做一做|问[:：]|想一想|不妨)",
    "过渡衔接": r"(接下来|下面(我们|来|请)|那么|于是|因此|所以|综上(所述)?|总之|可见|由此(可见|可得)|既然|反之|另一方面|首先.{0,8}其次|第一步|第二步|我们已经|我们知道|我们知道|前面(学|讲)|上文|下文|据此)",
    "归纳口诀": r"(口诀|记忆(法|口诀)?|巧记|速记|记住(这)?|三(步|看|种|大|条)|四(步|看|种|大|条)|五(步|看|种|大|条)|一(看|读|找|抓|辨)|归纳|总结|小结|提炼|建模|方法(总结|点拨|归纳)|规律|要点|技巧)",
    "警示易错": r"(注意|易错|特别(提醒|注意)|切记|千万不要|不要(把|将|误)|错误|误区|常见错误|区分|辨析|对比|陷阱|丢分|失分|扣分|重点|难点|关键|务必|牢记|区别|混淆)",
    "互动指令": r"(请(大家|同学们|你|各位)|小组(讨论|合作|交流)|讨论|回答|抢答|举手|同桌|动手|练习|完成(下面|下列|练习)?|填写|作答|看(图|表|课本|黑板|屏幕)|阅读(下面|下列|短文)|齐读|朗读|背诵|默写)",
    "讲评反馈": r"(讲评|得分|平均分|正确率|错误率|优秀|表扬|错题|失分|答题(情况|规范)|阅卷|评分|满分|参考答案|得分点|踩点|对答案|自查)",
    "思维方法": r"(思路|方法|分析[:：]|解题|步骤|模型|建构|框架|脉络|线索|主线|素养|能力|观念|思维|逻辑|推理|论证|证据|史料|实证)",
}

def clean(line):
    return re.sub(r"\s+", " ", line).strip()

def main():
    files = sorted(glob.glob(os.path.join(TXT, "*.txt")))
    per_cat = {c: {"count": 0, "files": set(), "examples": []} for c in CATS}
    per_subj = collections.defaultdict(lambda: collections.Counter())
    subj_files = collections.Counter()
    # 口诀专用：找短句+数字/对仗
    jué = []

    for f in files:
        rel = os.path.basename(f)
        subj = guess_subj(rel)
        subj_files[subj] += 1
        try:
            txt = open(f, encoding="utf-8", errors="replace").read()
        except Exception:
            continue
        # 按 slide/段落切
        lines = [clean(l) for l in txt.splitlines() if clean(l)]
        for ln in lines:
            if len(ln) > 120:  # 长句只取前段做样例
                sample = ln[:120]
            else:
                sample = ln
            for cat, pat in CATS.items():
                if re.search(pat, ln):
                    per_cat[cat]["count"] += 1
                    per_cat[cat]["files"].add(rel)
                    per_subj[subj][cat] += 1
                    if len(per_cat[cat]["examples"]) < 400:
                        per_cat[cat]["examples"].append({"subj": subj, "line": sample, "file": rel[:60]})
            # 口诀：短句、含数字列举或对仗（"一…二…三…"），且带教学词
            if re.search(r"(一[^，。]{2,6}[，、].{0,3}二[^，。]{2,6}|口诀|巧记|记忆)", ln) and len(ln) < 90:
                jué.append({"subj": subj, "line": sample})

    out = {
        "total_files": len(files),
        "subj_files": dict(subj_files.most_common()),
        "categories": {c: {"count": v["count"], "n_files": len(v["files"])} for c, v in per_cat.items()},
        "per_subj_cat": {s: dict(c) for s, c in per_subj.items()},
    }
    with open("analysis/full4/writecraft.json", "w", encoding="utf-8") as fp:
        json.dump(out, fp, ensure_ascii=False, indent=2)

    # digest：每类抽 12 句真实例句（去重）
    with open("analysis/full4/writecraft_digest.md", "w", encoding="utf-8") as fp:
        fp.write("# 写作工艺·语言标记挖掘（全 %d 份）\n\n" % len(files))
        fp.write("## 各维度命中（句数 / 覆盖文件数）\n\n")
        for c in CATS:
            fp.write("- **%s**：%d 句 / %d 文件\n" % (c, per_cat[c]["count"], len(per_cat[c]["files"])))
        fp.write("\n## 各维度真实例句\n")
        for c in CATS:
            fp.write("\n### %s\n" % c)
            seen = set()
            n = 0
            for e in per_cat[c]["examples"]:
                k = e["line"][:30]
                if k in seen:
                    continue
                seen.add(k)
                fp.write("- [%s] %s\n" % (e["subj"], e["line"]))
                n += 1
                if n >= 14:
                    break
        fp.write("\n## 口诀/记忆句样本\n")
        seen = set()
        n = 0
        for e in jué:
            if e["line"][:20] in seen:
                continue
            seen.add(e["line"][:20])
            fp.write("- [%s] %s\n" % (e["subj"], e["line"]))
            n += 1
            if n >= 40:
                break
        fp.write("\n## 各学科×维度分布\n\n| 学科 | " + " | ".join(CATS) + " | 文件数 |\n|---|---|" + "---|"*(len(CATS)+1) + "\n")
        for s, c in sorted(per_subj.items(), key=lambda x: -sum(x[1].values())):
            fp.write("| %s | " % s + " | ".join(str(c.get(k,0)) for k in CATS) + " | %d |\n" % subj_files[s])

    print("files:", len(files))
    print("categories:", json.dumps(out["categories"], ensure_ascii=False))
    print("subj_files:", json.dumps(out["subj_files"], ensure_ascii=False))
    print("jué samples:", len(jué))
    print("digest → analysis/full4/writecraft_digest.md")

if __name__ == "__main__":
    main()
