"""端到端验证：用官方 MCP 客户端连 stdio 服务，列工具并真实调用。"""
import asyncio
import json
import sys

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

PY = sys.executable


async def main():
    params = StdioServerParameters(
        command=PY,
        args=["-m", "jszz.server"],
        env={"PYTHONPATH": "/root/.hermes/workspace/jszz/server",
             "PATH": "/usr/local/bin:/usr/bin:/bin",
             "JSZZ_HOME": "/root/.hermes/workspace/jszz/.jszz-test"},
    )
    async with stdio_client(params) as (read, write):
        async with ClientSession(read, write) as session:
            await session.initialize()
            tools = await session.list_tools()
            print("TOOLS:", len(tools.tools))
            for t in tools.tools:
                print(" -", t.name)
            r = await session.call_tool("jszz_status", {})
            print("STATUS:", r.content[0].text[:400])

            r = await session.call_tool("bili_search", {"keyword": "匀变速直线运动", "page_size": 3})
            data = json.loads(r.content[0].text)
            print("BILI total:", data.get("total"), "first:", data["items"][0]["title"][:40])

            r = await session.call_tool("zxxk_search", {"keyword": "匀变速直线运动", "page_size": 3})
            data = json.loads(r.content[0].text)
            print("ZXXK ok:", data.get("ok"), "total:", data.get("total"))

            r = await session.call_tool("zxxk_preview_text", {"document_id": 60192049, "max_pages": 1})
            data = json.loads(r.content[0].text)
            print("PREVIEW ok:", data.get("ok"), "chars:", data.get("chars"))

            md = open("/root/.hermes/workspace/jszz/out/sample_xuean.md", encoding="utf-8").read()
            r = await session.call_tool("make_xuean", {"markdown": md})
            data = json.loads(r.content[0].text)
            print("XUEAN files:", data.get("files"))

            qs = json.dumps([
                {"stem": "关于匀变速直线运动，下列说法正确的是（　　）",
                 "options": ["速度随时间均匀变化", "加速度随时间均匀变化", "位移随时间均匀变化", "加速度方向与速度方向一定相同"],
                 "answer": "A", "score": 3, "type": "选择"},
                {"stem": "一物体做初速度为零的匀加速直线运动，第2 s内的位移为6 m，则加速度大小为____m/s²。",
                 "answer": "4", "score": 4, "type": "填空"},
                {"stem": "汽车以20 m/s行驶，刹车加速度大小为5 m/s²，求刹车后6 s内的位移。",
                 "answer": "40 m", "analysis": "刹车时间4 s，6 s内位移等于4 s内位移，x=v0t/2=40 m。",
                 "score": 12, "type": "解答"},
            ], ensure_ascii=False)
            r = await session.call_tool("assemble_paper", {
                "questions_json": qs, "title": "MCP端到端组卷测试", "total_score": 19})
            data = json.loads(r.content[0].text)
            print("ASSEMBLE:", data.get("files"), data.get("count"))


if __name__ == "__main__":
    asyncio.run(main())
