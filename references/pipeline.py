#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""金融创新作战地图 —— 自动补满流水线（每细分领域目标 5 条，去重不重复存）。

依赖：build.py（复用其 xlsx 读写、去重、HTML 组装）

用法：
  .venv/Scripts/python pipeline.py search            # 打印 20 个细分领域各自的搜索查询（交给人工/自动化去搜）
  .venv/Scripts/python pipeline.py status            # 检查 news.xlsx 每类当前条数 vs 目标 5，列出缺口
  .venv/Scripts/python pipeline.py fill <新新闻.json> # 去重并入、重建 HTML，并打印补满后每类状态
  .venv/Scripts/python pipeline.py rebuild           # 仅用 news.xlsx 重新生成 HTML

设计要点：
  - TAXONOMY 是唯一的"分类真理源"，地图图例/筛选/补满标准都以它为准。
  - 去重主键见 build.norm_key（链接归一化优先，标题兜底），已存在绝不重复存。
  - 真实新闻不足 5 条时脚本只报告缺口，绝不编造占位数据。
"""
import argparse, json, pathlib, sys
from collections import Counter

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))
import build

HERE = pathlib.Path(__file__).resolve().parent
XLSX = pathlib.Path.cwd() / "news.xlsx"   # 数据落在用户当前工作目录，不污染技能安装目录
TARGET = 5

# 20 个细分领域：(大类, 方向, 搜索查询)  —— 分类真理源
TAXONOMY = [
    # —— 银行（10 子类）——
    ("银行", "存款理财", "银行存款 理财 结构性存款 大额存单 创新 2026年8月"),
    ("银行", "公募私募", "银行 公募基金 私募 理财 代销 投顾 创新 2026年8月"),
    ("银行", "银保", "银行 银保 保险产品 代销 养老 年金 创新 2026年8月"),
    ("银行", "信托", "银行 信托 财富 家族信托 服务信托 创新 2026年8月"),
    ("银行", "个人信用贷款", "银行 个人信用贷款 消费贷 贴息 创新 2026年8月"),
    ("银行", "信用卡分期", "银行 信用卡 分期 权益 创新 2026年8月"),
    ("银行", "汽融", "银行 汽车金融 汽车分期 汽融 创新 2026年8月"),
    ("银行", "个人房贷", "银行 个人住房贷款 房贷 利率 创新 2026年8月"),
    ("银行", "个人房地产抵押贷款", "银行 个人房产抵押 经营贷 抵押贷 创新 2026年8月"),
    ("银行", "小微企业主贷款", "银行 小微 企业主 贷款 普惠 创新 2026年8月"),
    # —— 金融科技（5 子类）——
    ("金融科技", "产品创新", "银行 金融科技 产品创新 AI 大模型 2026年8月"),
    ("金融科技", "模式创新", "金融 科技 模式创新 开放银行 平台 2026年8月"),
    ("金融科技", "技术创新", "金融 科技 技术创新 区块链 大模型 风控 2026年8月"),
    ("金融科技", "需求走向", "金融 消费者需求 趋势 财富 信贷 2026年8月"),
    ("金融科技", "权益升级", "银行 权益 会员 升级 场景金融 2026年8月"),
    # —— 保险（5 险种）——
    ("保险", "寿险", "寿险 产品 创新 2026年8月 中国"),
    ("保险", "财险", "财产险 车险 农险 创新 2026年8月 中国"),
    ("保险", "健康险", "健康险 医疗险 创新 2026年8月 中国"),
    ("保险", "养老险", "养老保险 年金 个人养老金 创新 2026年8月 中国"),
    ("保险", "普惠保险", "普惠保险 惠民保 新市民 创新 2026年8月 中国"),
]

def _counts():
    news = build.xlsx_to_news(XLSX) if XLSX.exists() else []
    return news, Counter((n["category"], n["direction"]) for n in news)

def cmd_search():
    print("=== 20 个细分领域搜索查询（每日按此搜，去重并入即可补满）===\n")
    for i, (cat, dirn, q) in enumerate(TAXONOMY, 1):
        print(f"{i:2d}. [{cat} / {dirn}]\n    {q}\n")

def cmd_status():
    news, c = _counts()
    print(f"{'类别':<8}{'方向':<16}{'当前':>4}{'目标':>4}  状态")
    print("-" * 40)
    ok = 0
    for cat, dirn, _ in TAXONOMY:
        cnt = c.get((cat, dirn), 0)
        mark = "✓ 达标" if cnt >= TARGET else f"✗ 缺 {TARGET - cnt}"
        if cnt >= TARGET:
            ok += 1
        print(f"{cat:<10}{dirn:<18}{cnt:>4}{TARGET:>4}  {mark}")
    print("-" * 40)
    print(f"达标领域：{ok}/{len(TAXONOMY)}    总条目：{len(news)}")
    gaps = [(cat, dirn, TARGET - c.get((cat, dirn), 0)) for cat, dirn, _ in TAXONOMY if c.get((cat, dirn), 0) < TARGET]
    if gaps:
        print("缺口领域（需继续搜）：")
        for cat, dirn, need in gaps:
            print(f"   - {cat}/{dirn}  还需 {need} 条")
    else:
        print("全部细分领域已满足每类 5 条 ✓")

def cmd_fill(json_path):
    incoming = json.loads(pathlib.Path(json_path).read_text(encoding="utf-8"))
    existing = build.xlsx_to_news(XLSX) if XLSX.exists() else []
    before = len(existing) + len(incoming)
    merged = build.dedup(existing + incoming)      # existing 在前，已存在不重复
    skipped = before - len(merged)
    build.news_to_xlsx(merged, XLSX)
    build.build_html(merged, is_sample=False)
    print(f"并入传入 {len(incoming)} 条，跳过已存在 {skipped} 条，去重后合计 {len(merged)} 条 -> 已重写 news.xlsx 与 HTML\n")
    cmd_status()

def cmd_rebuild():
    build.build()

def main():
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="cmd", required=True)
    sub.add_parser("search")
    sub.add_parser("status")
    sub.add_parser("rebuild")
    f = sub.add_parser("fill"); f.add_argument("json")
    args = ap.parse_args()
    if args.cmd == "search": cmd_search()
    elif args.cmd == "status": cmd_status()
    elif args.cmd == "rebuild": cmd_rebuild()
    elif args.cmd == "fill": cmd_fill(args.json)

if __name__ == "__main__":
    main()
