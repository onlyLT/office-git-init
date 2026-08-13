# -*- coding: utf-8 -*-
"""把一个放着 Office 文档的文件夹，初始化成可协同的工作区。

用法（在目标文件夹里）：
    python <skill目录>/init_workspace.py            # 初始化当前目录
    python <skill目录>/init_workspace.py <目录>     # 初始化指定目录

做四件事，全部幂等（重复跑安全，不会覆盖已有配置）：
  1. git init（已是仓库则跳过）
  2. 装 .gittools/office2txt.py 并配好 textconv —— git diff 从此能看懂 docx/xlsx
  3. 写 .gitattributes（Office 文件挂 textconv）和 .gitignore（排除 ~$ 锁文件）
  4. 把现有文档提交为"基线版本"，作为日后比对和回滚的起点

初始化前会检查锁文件：有文档正开着就提示先关闭，避免把半截状态存成基线。
"""
import os
import shutil
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(sys.argv[1] if len(sys.argv) > 1 else ".")
OFFICE_EXT = (".docx", ".xlsx", ".doc", ".xls", ".pptx", ".ppt")
VERSION = "1.0"

CLAUDE_MD = """\
# 本工作区协作规程

<!-- 由 office-git-init skill v%s 生成。
     下方「协作纪律」「环境坑」可随 skill 更新刷新；
     「本项目专属约定」是你自己的内容，刷新时不会被动。 -->

本目录含 Office 文档（.docx/.xlsx），已建 git 库并配好内容级 diff。

## 协作纪律（改任何 docx/xlsx 都要走）

二进制文档协同有两个特有陷阱：**git 默认看不出内容差异**；**文档开着时写盘，
用户一保存就整体覆盖你的改动，且无任何报错**。所以：

### 1. 查锁——先分清"真占用"还是"僵尸锁"
`ls ~$*` 有输出说明存在锁文件。`~$` 常在 Office 异常退出后残留，别只凭它下结论：
用 `open(doc,"r+b")` 测独占打开，再看有无 Office 进程正开着这份文档。
两条都指向无人占用 → 残留锁，可改，但**别替用户删**，告知即可。
**确认有人开着 → 停下，让用户先关闭。**

### 2. 先 diff 再动手，并把用户的改动报告给他
```bash
git status --short
git diff -- "<文件名>"
```
**绝不拿自己早先的副本当"文件现在的样子"**——用户随时在改，副本在他碰文件那一刻
就过期了。它只能证明"我改过什么"，不能证明"文件现在是什么"。

### 3. 改
python-docx / openpyxl 定点改；格式复杂的表优先用 Office COM 原地改。
**改前先断言原值**，防止改错行：`assert "35 万元" in para.text`

### 4. 立刻提交
```bash
git add -A && git commit -m "谁改的 + 改了什么 + 为什么"
```
不提交的代价：用户的手工改动和你的改动混成一团，无法区分、无法回滚。
**这一步最容易被跳过**——实测中未加约束的助手 3/3 完成了任务却无一提交。

## 环境坑（Windows / Office）

- **WPS 常驻**时锁文件残留很普遍；备份别放会话临时目录，换会话即失效，用 git 提交代替备份。
- **用 Office COM 打开文件（即使 ReadOnly）本身就会回写文件**，"打开验证"后要重新提交快照。
- `.doc/.xls` 旧格式只能用 Office COM 读；加密的 .xls 用 Excel COM 传空密码打开。
- Word COM 的 `SaveAs2` 会卡死；COM 报 `RPC_E_CALL_REJECTED` 多半是 Office 弹了模态框。
- 改 xlsx 优先 Excel COM 原地改以保全格式；openpyxl 会重写整个文件。

## 常用命令

```bash
git diff -- "<文件>"                    # 看文档现在改了什么
git diff HEAD~1 HEAD -- "<文件>"        # 看某次改动
git show HEAD~2:"<文件>" > 旧版.docx     # 取历史版本另存（不覆盖当前）
```

## 本项目专属约定

<!-- 在这里写只属于本项目的规矩：数据口径、命名约定、哪些"异常"其实是正常的、
     专用校验脚本怎么跑等。这部分不会被 skill 更新覆盖。 -->

（待补充）
""" % VERSION

AGENTS_MD = """\
# AGENTS.md

本工作区的协作规程见 [CLAUDE.md](CLAUDE.md)。

要点：改任何 Office 文档前先查 `~$` 锁文件、先跑 `git diff` 看清他人改动，
改完立刻 `git commit`。详见 CLAUDE.md。
"""

GITATTRIBUTES = """\
# Office 文件是二进制：标为 binary 避免 git 误做行合并，
# 同时挂 textconv 驱动，让 git diff 显示内容级（段落/单元格）差异。
*.docx binary diff=office
*.xlsx binary diff=office
*.pptx binary diff=office
*.doc  binary diff=office
*.xls  binary diff=office
*.ppt  binary diff=office
*.pdf  binary
"""

GITIGNORE = """\
# Office 打开文档时生成的锁文件——存在即说明文档正被编辑，不入库
~$*

# Office 临时/自动恢复文件
*.tmp
*.asd
*.wbk
~WRL*.tmp

__pycache__/
*.pyc
"""


def sh(cmd, check=False):
    r = subprocess.run(cmd, shell=True, cwd=ROOT, capture_output=True,
                       text=True, encoding="utf-8", errors="replace")
    if check and r.returncode != 0:
        sys.exit("命令失败: %s\n%s" % (cmd, r.stderr))
    return r


def write_if_absent(name, content):
    p = os.path.join(ROOT, name)
    if os.path.exists(p):
        print("  · %s 已存在，保持不动" % name)
        return False
    with open(p, "w", encoding="utf-8", newline="\n") as f:
        f.write(content)
    print("  ✓ 已写入 %s" % name)
    return True


def main():
    if not os.path.isdir(ROOT):
        sys.exit("目录不存在: %s" % ROOT)
    print("初始化工作区: %s\n" % ROOT)

    # --- 0. 先查锁：有文档开着就别把半截状态存成基线 ---
    locks = [f for f in os.listdir(ROOT) if f.startswith("~$")]
    if locks:
        print("★ 检测到锁文件，说明有文档正开着（或上次异常退出残留）:")
        for f in locks:
            print("    %s" % f)
        print("  建议先关闭对应文档再初始化，避免把未保存完的状态存成基线。")
        print("  确认是残留锁可继续。\n")

    docs = [f for f in os.listdir(ROOT) if f.lower().endswith(OFFICE_EXT)
            and not f.startswith("~$")]
    print("发现 %d 个 Office 文档: %s\n" % (len(docs), "、".join(docs[:5]) or "无"))

    # --- 1. git ---
    if os.path.isdir(os.path.join(ROOT, ".git")):
        print("  · 已是 git 仓库，跳过 init")
    else:
        sh("git init -q", check=True)
        print("  ✓ git init")
    sh('git config core.quotepath false')          # 中文文件名正常显示

    # --- 2. textconv ---
    tools = os.path.join(ROOT, ".gittools")
    os.makedirs(tools, exist_ok=True)
    src = os.path.join(HERE, "office2txt.py")
    dst = os.path.join(tools, "office2txt.py")
    if os.path.exists(src):
        shutil.copyfile(src, dst)
        print("  ✓ 已装 .gittools/office2txt.py")
    else:
        print("  ★ 未找到 office2txt.py，textconv 将不可用")
    # 给不懂技术的同事看的说明书，随项目走
    rm = os.path.join(HERE, "版本管理说明.md")
    if not os.path.exists(rm):                       # 兼容旧版目录布局
        rm = os.path.join(HERE, "README.md")
    if os.path.exists(rm) and not os.path.exists(os.path.join(ROOT, "版本管理说明.md")):
        shutil.copyfile(rm, os.path.join(ROOT, "版本管理说明.md"))
        print("  ✓ 已装 版本管理说明.md（给同事看的入门说明）")
    sh('git config diff.office.textconv "python .gittools/office2txt.py"')
    sh('git config diff.office.cachetextconv false')
    print("  ✓ 已配 textconv（git diff 可显示文档内容差异）")

    # --- 3. 配置文件 ---
    write_if_absent(".gitattributes", GITATTRIBUTES)
    write_if_absent(".gitignore", GITIGNORE)

    # --- 3b. 规程文件：CLAUDE.md 无条件加载，规矩从此常驻本项目 ---
    if write_if_absent("CLAUDE.md", CLAUDE_MD):
        print("      ↑ 协作纪律已写入本项目，之后不必再依赖 skill")
    write_if_absent("AGENTS.md", AGENTS_MD)

    # --- 4. 基线提交 ---
    if sh("git log -1 --oneline").returncode != 0:      # 尚无任何提交
        sh("git add -A")
        sh('git -c core.safecrlf=false commit -q -m "基线版本：初始化协同工作区"')
        print("  ✓ 已提交基线版本")
    else:
        st = sh("git status --short").stdout.strip()
        if st:
            print("  · 已有提交历史；当前有未提交改动，未自动提交：")
            for line in st.splitlines()[:8]:
                print("      %s" % line)
            print("    确认无误后自行提交：git add -A && git commit -m \"...\"")
        else:
            print("  · 已有提交历史，工作区干净")

    # --- 自检 ---
    print("\n自检:")
    if docs:
        r = sh('git diff HEAD --stat -- "%s"' % docs[0])
        print("  textconv 可用性：", "✓" if r.returncode == 0 else "★ 异常")
    print("  下一步：改文档前先跑 git status --short 和 git diff，看清他人改动。")


if __name__ == "__main__":
    main()
