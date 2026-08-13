# office-git-init

> 让 git 看懂 Word / Excel —— 为「人 + AI 协同编辑 Office 文档」一键搭好带内容级 diff 的版本管理工作区。
>
> An [Agent Skill](https://docs.claude.com/en/docs/claude-code/skills) that turns any folder of Office documents into a git workspace with content-level diffs, built for human–AI co-editing.

## 为什么需要它

让 AI 助手帮你改 `.docx` / `.xlsx`，有两个代码文件上不存在的陷阱：

**1. 看不见差异。** git 对二进制文档只会说：

```
$ git diff
 季度总结.docx | Bin 35533 -> 35565 bytes
```

改了哪一段、哪个单元格？不知道。助手说"改好了三处"，你只能选择相信。

**2. 静默覆盖。** 文档开在 Word/WPS 里时，助手写盘后你一按保存，内存里的旧版本会**整体覆盖**磁盘文件——几百处改动瞬间归零，**没有任何报错**。

跑一次本 skill 的初始化脚本之后：

```
$ git diff -- 季度总结.docx
-收入 100 万元，支出 80 万元。
+收入 120 万元，支出 90 万元。

$ git diff -- 明细表.xlsx
-  A2=收入 | B2=100
+  A2=收入 | B2=150
```

逐段、逐单元格看清每一处改动，脚注和公式也在监控范围内（公式被改成常量能直接看出来）。AI 的每次修改都有提交记录，可核查、可回滚。

## 它做了什么

在目标文件夹里跑一次 `init_workspace.py`（幂等，重复跑安全），它会：

1. **`git init`** —— 已是仓库则跳过；
2. **安装 textconv 驱动** —— 把 `office2txt.py` 装进 `.gittools/` 并配好 `diff.office.textconv`，`git diff` 从此显示段落 / 单元格 / 脚注级差异；
3. **写好 `.gitattributes` 和 `.gitignore`** —— Office 文件挂 textconv 并标为 binary（防止 git 误做行合并），`~$` 锁文件不入库；
4. **生成协作规程** —— 向目标项目写入 `CLAUDE.md` / `AGENTS.md`（AI 助手每次会话自动加载的协作纪律：查锁 → 先 diff → 定点改 → 立刻提交）和《版本管理说明.md》（给不懂编程的同事看的四条命令入门）；
5. **提交基线版本** —— 作为日后比对和回滚的起点。初始化前先查锁文件，避免把编辑到一半的状态存成基线。

之后的日常协作**不依赖本 skill**：纪律已常驻目标项目的 `CLAUDE.md`，命令全是标准 git。

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

### 不用 AI agent，独立使用

脚本不依赖任何 AI 工具，可以直接跑：

```bash
cd <放着 Office 文档的文件夹>
python <本仓库路径>/init_workspace.py
```

**依赖**：git、Python 3，以及 `pip install python-docx openpyxl lxml`（textconv 解析 docx/xlsx 用）。

## 协作纪律（初始化后写入目标项目）

这是本 skill 的核心——工具只解决"看得见"，纪律解决"不出事"：

| 步骤 | 做什么 | 防的是什么 |
|---|---|---|
| 1. 查锁 | 见 `~$` 锁文件先判定是"真占用"还是"僵尸锁"，有人开着就停下 | 写盘后被用户的保存动作整体覆盖 |
| 2. 先 diff | `git diff` 看清对方改了什么，绝不拿早先的副本当现状 | 拿旧版覆盖对方的新改动 |
| 3. 定点改 | python-docx / openpyxl / Office COM 修改，改前断言原值 | 改错行、破坏勾稽关系 |
| 4. 立刻提交 | 每个节点 `git commit`，写明改了什么、依据是什么 | 两人的改动混成一团，无法区分回滚 |

`SKILL.md` 里还沉淀了一批实测踩出来的 Windows / Office 环境坑：WPS 僵尸锁的判定方法、Office COM"只读打开也会回写文件"、加密 `.xls` 的空密码处理、`RPC_E_CALL_REJECTED` 的真实原因（模态框）等。

## 仓库结构

```
office-git-init/
├── SKILL.md            # skill 本体：触发条件、工作流、Windows/Office 踩坑清单
├── init_workspace.py   # 一键初始化脚本（幂等）
├── office2txt.py       # textconv 驱动：docx/xlsx → 纯文本（段落/表格/脚注/公式）
├── 版本管理说明.md      # 初始化时装进目标项目，给不懂编程的同事看的入门说明
└── README.md           # 本文件
```

## 常见问题

**git 会把我的文档改坏吗？** 不会。仓库存的是完整原始文件，一个字节不差；`git diff` 显示的文字只是转换出来看的，不影响存储。

**能像代码那样自动合并两个人的修改吗？** 不能。Office 文档是二进制，git 合并不了，所以"同一时刻只有一方在改"这条纪律比在代码项目里更重要。

**`.doc` / `.xls` 旧格式支持吗？** 会纳入版本管理，但 diff 只能显示指纹（md5 + 大小），判断"变没变"；内容级 diff 仅支持 `.docx` / `.xlsx`。

## License

[MIT](LICENSE)
