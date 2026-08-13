---
name: office-git-init
description: Use when first setting up a folder that contains Word/Excel/PowerPoint files (.docx/.xlsx/.doc/.xls/.pptx) for human-AI collaboration — a new document project, a folder with no git repo, or one where git diff shows only "Bin 35533 -> 35565 bytes" instead of what changed inside the document.
---

# 与人协同编辑 Office 文档

> **本 skill 是一次性的初始化工具。** 跑完 `init_workspace.py`，协作纪律就写进了
> 目标项目的 `CLAUDE.md`（无条件加载，不依赖本 skill 触发），日常工作照那份走即可，
> 不必再回来读这里。下面的内容是初始化会装进项目的东西，以及为什么这么装。

## Overview

二进制办公文档（.docx/.xlsx）的协同有两个特有陷阱，代码文件上不存在：

1. **看不见差异**——git 只会说 `Bin 35533 -> 35565 bytes`，说不出改了哪一段、哪个单元格。
2. **静默覆盖**——文档开在 Word/WPS 里时，你写盘后用户一保存，内存里的旧版本会**整体覆盖**磁盘版本，几百处改动瞬间归零，**没有任何报错**。

核心原则：**磁盘上的文件是唯一事实；改前先看差异，改后立刻留痕。**

## When to Use

- 用户或同事也在编辑同一个 Word/Excel 文件
- 用户说"我改了一些""你帮我看看他改了什么"
- 目录里有 `~$xxx.docx` 锁文件
- 需要知道二进制文档内部改了什么

**不适用**：纯文本/代码文件（git 原生可 diff）；只读不改的场合。

## 工作区初始化（每个文件夹做一次）

**接手任何放着 Office 文档的文件夹，第一件事就是跑这条。** 没有它，后面的
"先 diff"根本无从谈起——git 对二进制文档只会说 `Bin 35533 -> 35565 bytes`。

```bash
cd <目标文件夹>
python <本skill目录>/init_workspace.py
```

幂等，重复跑安全。它会：git init → 装 `.gittools/office2txt.py` 并配好 textconv
→ 写 `.gitattributes`（Office 文件挂 textconv）和 `.gitignore`（排除 `~$` 锁文件）
→ 把现有文档提交为"基线版本"。初始化前先查锁，有文档开着会提示，避免把
半截状态存成基线。

配好之后的效果：

```
$ git diff -- 季度总结.docx
-收入 100 万元，支出 80 万元。
+收入 120 万元，支出 90 万元。

$ git diff -- 明细表.xlsx
-  A2=收入 | B2=100
+  A2=收入 | B2=150
```

**已有历史的文件夹**：脚本不会自动提交已存在的改动，只提示你确认后自行提交——
避免把别人未完成的活当成基线。

## 每次改动的四步

**这四步是一个整体，缺任何一步都会在协同时出事。**

### 1. 查锁——先分清"真占用"还是"僵尸锁"
```bash
ls ~$* 2>/dev/null      # 有输出 = 可能正开着
```
`~$` 文件常在 Office 异常退出后残留，见到它先别停，做两项判定：

```python
try:                                   # ① 能独占打开 = 无进程持有
    open(doc, "r+b").close(); free = True
except PermissionError:
    free = False
```
② 查 Office 进程实际开着哪些文件（窗口标题或命令行）。

- **两项都表明无人占用** → 是残留锁，可以改；改完提醒用户删掉它。
- **确有进程开着这份文档** → **停下，请用户先关闭**，不要硬写。硬写的两种下场：
  报 `Device or resource busy`（好情况，至少你知道失败了），或写成功但随后被用户
  的保存动作整体覆盖（坏情况，无声无息，几百处改动瞬间归零）。

### 2. 先 diff 再动手，并把用户的改动报告给他
```bash
git status --short
git diff -- "<文件名>"
```
**绝不拿自己早先的副本当"文件现在的样子"。** 用户随时在改，你 scratchpad 里的
副本在他碰文件那一刻就过期了——它只能证明"我改过什么"，不能证明"文件现在是什么"。

### 3. 改
用 python-docx / openpyxl 做定点修改；格式复杂的表格优先用 Office COM 原地改。
**改之前先断言原值**，防止改错行：
```python
assert "35 万元" in para.text, "原值不符，已被他人改动，停止"
```

### 4. 立刻提交——这一步最容易被跳过
```bash
git add -A && git commit -m "谁改的 + 改了什么 + 为什么"
```
不提交的代价：用户的手工改动和你的改动混成一团，无法区分、无法回滚到中间点。
基线测试中 3/3 的代理都完成了任务却**没有一个提交**，正是这个坑。

## Quick Reference

| 你要做的事 | 命令 |
|---|---|
| 看文档现在什么样 | `git diff -- "<文件>"`（已配 textconv） |
| 看某次改动 | `git diff HEAD~1 HEAD -- "<文件>"` |
| 取出历史版本另存（不覆盖当前） | `git show HEAD~2:"<文件>" > 旧版.docx` |
| 回滚到上一版（会覆盖，慎用） | `git checkout HEAD~1 -- "<文件>"` |
| 文件被锁又必须记录版本 | `git hash-object -w` + `update-index --cacheinfo` + `commit-tree` |

## Common Mistakes

| 做法 | 后果 | 正确做法 |
|---|---|---|
| 拿早先的副本当现状判断依据 | 把用户的改动误判成"数据丢失"，白排查半天 | 只认 `git diff` 和磁盘文件 |
| 把备份存进会话临时目录 | 换个会话就没了，等于没备份（基线测试中 2/3 犯此错） | 备份进项目目录，或干脆用 `git commit` 代替备份 |
| 文档开着照样写盘 | 用户一保存，你的改动全没，无报错 | 先让用户关闭 |
| 改完不提交 | 两人的改动混在一起，无法区分和回滚 | 每个节点提交，信息写清依据 |
| 用 Office COM 打开"只读"看一眼就以为没变 | COM 打开即回写（实测脚注 id 被重排） | 打开验证后重新提交快照 |
| 直接覆盖单元格公式 | 破坏表内勾稽 | 先读 `data_only=False` 看清是值还是公式 |
| 靠脚本"全绿"就宣布没问题 | 脚本只验它写死的那些项 | 脚本 + `git diff` 逐字看，两者都要 |

## Windows / Office 环境坑

- **WPS 常驻**：十几个 `wps.exe` 是常态，锁文件残留很普遍。残留锁（正文已移走/
  已关闭）可删，删前用独占打开测一下：`open(doc,'r+b')` 成功即无进程占用。
- **`.doc/.xls` 旧格式**：python-docx/openpyxl 读不了，只能用 Office COM。
- **加密的 .xls**：`xlrd` 报 `Workbook is encrypted`，用 Excel COM 打开并传空密码
  （`Password=""`）避免弹模态框卡死。
- **Word COM 的 `SaveAs2` 会卡死**：改用直接读 `Paragraphs.Item(i).Range/Format`。
- **COM 调用返回 `RPC_E_CALL_REJECTED`**：多半是 Office 弹了模态框（如首次运行的
  "接受许可协议"）。枚举窗口找出来让用户点掉，不要盲目重试。
- **保全格式**：改 xlsx 优先 Excel COM 原地改；openpyxl 会重写整个文件，仅在确认
  工作簿是纯单元格表格（无图表/透视表/图片）时使用。

## Red Flags — 出现这些念头就停下

- "我记得这个文件是……" → 你不记得，去 `git diff`
- "先改完再一起提交" → 现在就提交，混在一起就分不开了
- "锁文件应该是残留吧" → 测一下再说，猜错的代价是用户的活白干
- "改动很小，不用留痕" → 小改动混进大改动里，出事时最难查
- "脚本跑过了，没问题" → 脚本只验它认识的项，diff 还得看
