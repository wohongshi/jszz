# B站素材手册

## 1. 能做什么

| 需求 | 工具 | 说明 |
|---|---|---|
| 找教学视频 | `bili_search` | 支持排序、时长筛选 |
| 看视频详情/分P | `bili_video_info` | 拿 cid 才能取字幕 |
| 取字幕文本 | `bili_subtitle` | 带 `[mm:ss]`，可当逐字稿 |
| 下音频（mp3） | `bili_download(audio_only=True)` | 听力、背景音素材 |
| 下视频（mp4） | `bili_download(audio_only=False, quality="1080")` | 需要 yt-dlp + ffmpeg |
| 取音频直链 | `bili_audio_stream` | 不落盘，交给播放器/转写 |

## 2. 搜索参数

```text
order:    totalrank 综合 | click 播放多 | pubdate 最新 | dm 弹幕多 | stow 收藏多
duration: 0 全部 | 1 <10min | 2 10–30min | 3 30–60min | 4 >60min
```

备课场景常用：`order="click", duration=3`（30–60 分钟的完整课），
或 `order="pubdate"` 找最新的一轮复习。

## 3. 签名机制（为什么不用第三方 SDK）

B 站搜索接口要求 WBI 签名：

1. `GET /x/web-interface/nav` 拿 `wbi_img.img_url` / `sub_url`，取文件名（去掉扩展名）
2. 拼成 64 字符，按固定 64 位置表重排，取前 32 位 = `mixin_key`
3. 参数加 `wts=时间戳`，按 key 排序 urlencode（去掉 `!'()*`），
   `w_rid = md5(query + mixin_key)`

实现在 `server/jszz/bilibili.py`，约 40 行，不依赖 `bilibili-api`。
nav 接口拿不到时有一组长期有效的兜底 key（匿名搜索够用）。

## 4. Cookie 与清晰度

- 未登录：能搜、能拿 360P/480P，1080P 多数可下，AI 字幕拿不到
- 登录（SESSDATA）：1080P+/4K、AI 字幕、会员内容
- 配置方式：WebUI「B站账号」粘贴 `SESSDATA=...`，或
  `jszz login bilibili "SESSDATA=...; bili_jct=..."`
- cookie 存在 `~/.jszz/cookies/bilibili.txt`（权限 600）

## 5. 下载

`bili_download` 走 `yt-dlp`，命令等价于：

```bash
yt-dlp "https://www.bilibili.com/video/BVxxxx" \
  -f "bestvideo[height<=1080]+bestaudio/best[height<=1080]" \
  --add-header "Cookie:SESSDATA=..." \
  --add-header "Referer:https://www.bilibili.com/" \
  -o "%(title).120B.%(ext)s" --print after_move:filepath
```

音频：`-x --audio-format mp3 --audio-quality 0`。

## 6. 课堂使用建议

- 一条 40 分钟视频只截 2–4 分钟片段，用 `--download-sections` 可只截片段（yt-dlp 参数）
- 听力素材统一转 128kbps mp3，避免教室音响破音
- 字幕稿可以喂给 `make_xuean`，把视频内容变成学案里的「情境导入」
- 引用视频时在课件页脚标注 UP 主与 BV 号，尊重原作者
