# office-git-init

> 给 Word / Excel 文档装上"修订记录 + 后悔药"。在文件夹里跑一次，之后 AI 帮你改文档的每一处都**看得见、退得回、盖不掉**。
>
> An [Agent Skill](https://docs.claude.com/en/docs/claude-code/skills) that makes AI edits to Office documents auditable, reversible, and overwrite-safe — by turning any folder of `.docx`/`.xlsx` files into a git workspace with content-level diffs.

## 这是什么

一个**一次性初始化工具**：在任何放着 Office 文档的文件夹里跑一次脚本，这个文件夹从此具备三样能力——

- **看得见**：AI（或同事）改了哪一段、哪个单元格，一条命令全部列出来；
- **退得回**：每一步都有存档，改坏了随时回到任何一版；
- **盖不掉**：防"文档开着时被静默覆盖"的检查和协作纪律，常驻项目自动生效。

初始化之后日常协作**不依赖本 skill**：用的全是标准 git 命令，规矩写进了项目的 `CLAUDE.md` / `AGENTS.md` 由 agent 自动遵守，不懂编程的同事看一页《版本管理说明》也能参与。

## 它解决的三个真实问题

### 1. AI 说"改好了"，你没法核对

你让 AI 改一份几十页的报告、上千行的表格。它说"改了三处"——是不是真的只改了三处？有没有顺手动了别的数字？逐页人工核对不现实，不核对就只能选择相信。

git 本来帮不上忙，对 Word/Excel 它只会说：

```
$ git diff
 季度总结.docx | Bin 35533 -> 35565 bytes
```

初始化之后，同一条命令变成：

```
$ git diff -- 季度总结.docx
-收入 100 万元，支出 80 万元。
+收入 120 万元，支出 90 万元。

$ git diff -- 明细表.xlsx
-  A2=收入 | B2=100
+  A2=收入 | B2=150
```

正文逐段、表格逐单元格，连**脚注**和**公式**都在监控范围内——公式被悄悄改成写死的数值，一眼就能看出来。AI 的工作从"只能相信"变成"可以核查"：它说改了三处，你跑一下 `git diff` 就知道是不是真的只有三处。

### 2. 改坏了，回不去

磁盘上的 Word/Excel 文件没有版本历史，覆盖保存就是永别。"改之前这个数是多少来着？""上周报出去的是哪一版？"——只能翻邮件、翻微信找旧附件，靠《报告-最终版(3).docx》这种命名自救。

初始化之后，每次改动都有存档：

- 任何一版都能**原样取出来**（不动当前文件，另存一份对照）；
- 确定要回退时一条命令回到上一版；
- 每次存档写明**改了什么、依据是什么**，天然形成一份可追溯的工作台账——事后说得清"这个数字为什么从 1200 变成 1350"。

### 3. 你和 AI 互相覆盖，而且没有任何提示

这是 Office 文档协同特有的事故：AI 在磁盘上改完了，你这边 Word 还开着几分钟前的旧版本——你一按保存，它刚才的几百处改动**瞬间归零，不报错、不提示**。反过来，AI 拿着早先的旧副本去改，也会盖掉你刚改的内容。等发现时可能已经过去好几天。

初始化会把一套四步纪律写进项目（agent 每次会话自动加载，你随时可以用 `git diff` / `git log` 检查它有没有照做）：

| 步骤 | 做什么 | 防的是什么 |
|---|---|---|
| 1. 查锁 | 见 `~$` 锁文件先判定"真占用"还是"僵尸锁"，有人开着就停下 | 写盘后被用户的保存动作整体覆盖 |
| 2. 先 diff | 看清对方改了什么，绝不拿早先的副本当"文件现在的样子" | 拿旧版覆盖对方的新改动 |
| 3. 定点改 | python-docx / openpyxl / Office COM 修改，改前断言原值 | 改错行、破坏表内勾稽 |
| 4. 立刻提交 | 每个节点存档，写明改了什么、依据是什么 | 两人的改动混成一团，无法区分回滚 |

就算事故还是发生了——只要存过档，被覆盖的内容**一定找得回来**。

## 它是怎么做到的

在目标文件夹里跑一次 `init_workspace.py`（幂等，重复跑安全）：

```bash
cd <放着 Office 文档的文件夹>
python <本仓库路径>/init_workspace.py
```

它做五件事：

1. **`git init`** —— 已是仓库则跳过；
2. **装内容级 diff** —— 把 `office2txt.py` 装进 `.gittools/` 并配好 git 的 textconv 机制，`git diff` 从此显示段落 / 单元格 / 脚注级差异；
3. **写好 `.gitattributes` 和 `.gitignore`** —— Office 文件标为 binary（防止 git 误做行合并），`~$` 锁文件不入库；
4. **生成协作规程** —— 向目标项目写入 `CLAUDE.md` / `AGENTS.md`（上面那套四步纪律 + Windows/Office 实测踩坑清单）和《版本管理说明.md》（给不懂编程的同事看的四条命令入门）；
5. **提交基线版本** —— 作为日后比对和回滚的起点。初始化前先查锁文件，避免把编辑到一半的状态存成基线。

原理一句话：git 的 textconv 机制把 docx/xlsx 转成纯文本再比对。**仓库里存的仍是完整原始文件，一个字节不差**——转换出来的文字只用于显示，不影响存储，随时能取出和原件完全相同的历史版本。

## 安装

本 skill 采用通用的 [Agent Skills](https://docs.claude.com/en/docs/claude-code/skills) 格式（`SKILL.md` + 脚本），凡是支持这个格式的 agent 都能用。**安装方式只有一句话：把本仓库克隆到你的 agent 的 skills 目录。**

```bash
git clone https://github.com/onlyLT/office-git-init.git <你的 agent 的 skills 目录>/office-git-init
```

各家 agent 的 skills 目录：

| Agent | 个人级（所有项目可用） | 项目级（随代码库走） |
|---|---|---|
| Claude Code | `~/.claude/skills/office-git-init` | `<项目>/.claude/skills/office-git-init` |
| Codex CLI | `~/.codex/skills/office-git-init` | — |
| WorkBuddy | `~/.workbuddy/skills/office-git-init` | `<项目>/.workbuddy/skills/office-git-init` |
| 其他 | 参考其文档中的 skills 目录，放入即可 | 同左 |

例如安装到 Claude Code：

```bash
git clone https://github.com/onlyLT/office-git-init.git ~/.claude/skills/office-git-init
```

安装后无需手动触发——当 agent 接手一个放着 Office 文档、还没配版本管理的文件夹时，会自动匹配到本 skill 并提议初始化。也可以直接说"用 office-git-init 初始化这个文件夹"。（Codex 新增 skill 后若 `/skills` 里没出现，重启一次即可；WorkBuddy 需重启后生效。）

由于初始化时同时生成 `AGENTS.md` 和 `CLAUDE.md`，**建好的工作区本身与 agent 无关**——不管当初用哪个 agent 初始化的，其他遵循 AGENTS.md 约定的助手接手时同样会加载协作纪律。


## 仓库结构

```
office-git-init/
├── SKILL.md            # skill 本体：触发条件、四步纪律、Windows/Office 踩坑清单
├── init_workspace.py   # 一键初始化脚本（幂等）
├── office2txt.py       # textconv 驱动：docx/xlsx → 纯文本（段落/表格/脚注/公式）
├── 版本管理说明.md      # 初始化时装进目标项目，给不懂编程的同事看的入门说明
└── README.md           # 本文件
```

`SKILL.md` 里还沉淀了一批实测踩出来的 Windows / Office 环境坑：WPS 僵尸锁的判定方法、Office COM"只读打开也会回写文件"、加密 `.xls` 的空密码处理、`RPC_E_CALL_REJECTED` 的真实原因（模态框）等。

## 常见问题

**git 会把我的文档改坏吗？** 不会。仓库存的是完整原始文件，一个字节不差；`git diff` 显示的文字只是转换出来看的，不影响存储。

**能像代码那样自动合并两个人的修改吗？** 不能。Office 文档是二进制，git 合并不了，所以"同一时刻只有一方在改"这条纪律比在代码项目里更重要——代码冲突了还能合，文档冲突了只能二选一。

**`.doc` / `.xls` 旧格式支持吗？** 会纳入版本管理，但 diff 只能显示指纹（md5 + 大小），判断"变没变"；内容级 diff 仅支持 `.docx` / `.xlsx`。

**不懂 git 的同事怎么办？** 初始化时会往项目里放一份《版本管理说明.md》，只讲日常要用的四条命令（看状态、看改动、存档、看历史），不需要懂编程。

## License

[MIT](LICENSE)
