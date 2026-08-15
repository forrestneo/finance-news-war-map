"""Supabase 读写（配置驱动）。
使用前：
  1) 在 config.toml 填入 url / key / table
  2) pip install supabase（本 venv 已装）
  3) 数据库中建表，列：time,category,direction,company,city,lng,lat,title,link,event,why,innovation,learn
"""
import tomllib, pathlib, json

CFG = pathlib.Path(__file__).with_name("config.toml")

def load_cfg():
    if not CFG.exists():
        raise FileNotFoundError("缺少 config.toml，请复制 config.example.toml 并填写 Supabase 凭据")
    return tomllib.loads(CFG.read_text(encoding="utf-8"))["supabase"]

def _client():
    try:
        from supabase import create_client
    except ImportError:
        raise RuntimeError("未安装 supabase，请运行：.venv/Scripts/pip install supabase")
    c = load_cfg()
    return create_client(c["url"], c["key"]), c["table"]

def pull():
    """从 Supabase 拉取全部新闻，返回 list[dict]（含 coord）。"""
    client, table = _client()
    rows = client.table(table).select("*").order("time", desc=True).execute().data
    out = []
    for r in rows:
        out.append({
            "time": r.get("time"), "category": r.get("category"), "direction": r.get("direction"),
            "company": r.get("company"), "city": r.get("city"),
            "coord": [float(r.get("lng")), float(r.get("lat"))],
            "title": r.get("title"), "link": r.get("link"),
            "event": r.get("event"), "why": r.get("why"),
            "innovation": r.get("innovation"), "learn": r.get("learn"),
        })
    return out

def push(news):
    """把 news 列表写入 Supabase（按 title 做 upsert 需表有唯一约束；否则用 insert）。"""
    client, table = _client()
    rows = [{
        "time": n["time"], "category": n["category"], "direction": n["direction"],
        "company": n["company"], "city": n["city"],
        "lng": n["coord"][0], "lat": n["coord"][1],
        "title": n["title"], "link": n["link"],
        "event": n["event"], "why": n["why"],
        "innovation": n["innovation"], "learn": n["learn"],
    } for n in news]
    # 优先 upsert；若表无唯一键则退化为 insert
    try:
        client.table(table).upsert(rows, on_conflict="title").execute()
    except Exception as e:
        print("upsert 失败，改为 insert：", e)
        client.table(table).insert(rows).execute()
    print(f"已推送 {len(rows)} 条到 Supabase 表 {table}")
