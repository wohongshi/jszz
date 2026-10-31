#!/usr/bin/env python3
"""平衡采样各科写作工艺 + 挖「写作规则类」元内容（怎么答/怎么赋分/口诀/规范表述）。"""
import json, re, glob, os, collections, random

TXT = "analysis/full4/txt"
SUBJ = {
    "语文": ["语文","文言","古诗","作文","《","师说","劝学","赤壁","烛之武","种树","石钟山","默写","诗歌","小说","散文","记叙","议论文"],
    "数学": ["数学","函数","导数","数列","三角","向量","概率","立体几何","解析几何","不等式","集合","圆锥","导数"],
    "英语": ["英语","Unit","B1U","B2U","B3U","完形","语法","时态","从句","晨诵","晚写","It句型","应用文","读后续写"],
    "物理": ["物理","动量","受力","电场","磁场","振动","波","光学","原子","电路","动能","冲量","守恒定律"],
    "化学": ["化学","反应速率","离子","有机","氧化","催化","电极","电池","平衡","电解","工艺流程","金属","非金属","溶液"],
    "生物": ["生物","细胞","基因","遗传","光合","呼吸","生态","神经","激素","酶","抗体","免疫","DNA"],
    "政治": ["政治","经济","哲学","文化","法治","唯物","辩证","时政","价值观"],
    "历史": ["历史","朝代","近代","古代","改革","革命","史官","纲要","史料"],
    "地理": ["地理","气候","地形","大气","水文","洋流","地质","人口","城市","农业","工业","等高线","热力环流","气压带"],
}
def guess(rel):
    for s, ks in SUBJ.items():
        for k in ks:
            if k.lower() in rel.lower(): return s
    return "未分类"

def clean(s): return re.sub(r"\s+"," ",s).strip()

# 写作规则元内容（教你怎么写/怎么答/怎么赋分）
META = {
    "赋分/采分点": r"(采分|得分点|赋分|给分|得分关键词|按点作答|分[行点]书写|踩点|独立给分|要点独立)",
    "禁口语/规范表述": r"(不得分|禁用|口语|规范表述|错误写法|标准表述|专业术语|书面语|零分|不能写成)",
    "口诀对仗": r"([，、][^，。]{2,8}[，、][^，。]{2,8}[。！]|口诀|巧记|记忆口诀|速记|三字经|顺口溜)",
    "短句/信息点": r"(短句|信息点|主谓宾|写足|一句话.{0,6}得分|要点)",
    "双栏对照": r"(错误写法.{0,4}规范|左边.{0,4}右边|对比.{0,4}(正确|错误)|√|×|✔|✘)",
}

# 各科写作气质：学习目标/素养/考情 的措辞
CRAFT = {
    "学习目标动词": r"(了解|理解|掌握|运用|能(提取|利用|判断|说明|解释|分析|设计)|会用|认识|体会|培养|形成)",
    "素养四维": r"(人地协调观|区域认知|综合思维|地理实践力|证据推理|宏观辨识|微观探析|变化观念|平衡思想|科学探究|生命观念|科学思维|社会责任|史料实证|历史解释|唯物史观|家国情怀|政治认同|科学精神|法治意识|公共参与)",
    "考情分析": r"(命题(趋势|规律|方向|情境|角度)|考情|考点提取|考查(重点|方式|角度)|高频考点|五年(真题|高频)|考向|命题统计)",
    "设问句式": r"(为什么|如何|怎样|怎么(样|做)|是什么|试(问|着)|请(思考|分析|说明|回答)|想一想|思考[:：]|讨论)",
    "三步/四看": r"([三四五]步法|[三四五]看|[三四五]大|[三四五]种|口诀[:：]|一.{0,3}二.{0,3}三)",
}

def main():
    files = sorted(glob.glob(os.path.join(TXT,"*.txt")))
    by = collections.defaultdict(list)
    for f in files:
        by[guess(os.path.basename(f))].append(f)

    out = {"meta": collections.defaultdict(lambda: collections.defaultdict(list)),
           "craft": collections.defaultdict(lambda: collections.defaultdict(list)),
           "subj_n": {s: len(v) for s, v in by.items()}}

    for subj, flist in by.items():
        random.seed(42)
        sample = flist if len(flist) <= 60 else random.sample(flist, 60)
        for f in sample:
            try: txt = open(f, encoding="utf-8", errors="replace").read()
            except Exception: continue
            for ln in txt.splitlines():
                ln = clean(ln)
                if not ln or len(ln) > 140: ln = ln[:140]
                for cat, pat in META.items():
                    if re.search(pat, ln) and len(out["meta"][subj][cat]) < 60:
                        out["meta"][subj][cat].append(ln)
                for cat, pat in CRAFT.items():
                    if re.search(pat, ln) and len(out["craft"][subj][cat]) < 60:
                        out["craft"][subj][cat].append(ln)

    # 汇总 markdown
    with open("analysis/full4/writecraft2.md","w",encoding="utf-8") as fp:
        fp.write("# 写作工艺·跨科平衡采样 + 写作规则元内容\n\n")
        fp.write("## 各科样本文件数\n" + json.dumps(out["subj_n"], ensure_ascii=False) + "\n")
        fp.write("\n## 写作规则元内容（怎么写/怎么答/怎么赋分）\n")
        for cat in META:
            fp.write("\n### %s\n" % cat)
            for subj in ["语文","数学","英语","物理","化学","生物","政治","历史","地理","未分类"]:
                ex = out["meta"][subj][cat][:4]
                if not ex: continue
                fp.write("**%s**：\n" % subj)
                for e in ex: fp.write("- %s\n" % e)
        fp.write("\n## 各科写作气质（措辞风格）\n")
        for cat in CRAFT:
            fp.write("\n### %s\n" % cat)
            for subj in ["语文","数学","英语","物理","化学","生物","政治","历史","地理","未分类"]:
                ex = out["craft"][subj][cat][:4]
                if not ex: continue
                fp.write("**%s**：\n" % subj)
                for e in ex: fp.write("- %s\n" % e)

    with open("analysis/full4/writecraft2.json","w",encoding="utf-8") as fp:
        json.dump({k: {s: dict(c) for s, c in v.items()} for k, v in out.items() if k != "subj_n"} | {"subj_n": out["subj_n"]}, fp, ensure_ascii=False, indent=2)
    print("subj_n:", json.dumps(out["subj_n"], ensure_ascii=False))
    print("→ analysis/full4/writecraft2.md")

if __name__ == "__main__":
    main()
