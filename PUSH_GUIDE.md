# 首次推送到 GitHub 操作说明

本仓库的文件已准备齐全，达到可直接推送状态。当前环境还有两项前置条件需要你先处理，下面按顺序列出。

## 0. 当前环境状况（已实测）

| 项目 | 状态 |
| --- | --- |
| `git` | 已安装，版本 2.55.0 |
| `git config user.name` / `user.email` | **均未配置**，直接 commit 会失败 |
| `gh` CLI | 未安装 |
| 到 github.com 的网络 | 已实测可通（对 `octocat/Hello-World` 成功返回引用），代理已放行 |

因此需要先配置提交身份、准备认证方式，并确认网络可访问 GitHub。

## 1. 在 GitHub 网页端新建仓库

打开 `https://github.com/ivy92888-glitch`，新建仓库：

- Repository name：`sheet-keyweave`
- 可见性：Public
- **不要**勾选 `Add a README file` / `Add .gitignore` / `Choose a license`（本地已备齐，勾选会造成首次推送冲突）

创建后 GitHub 会给出仓库地址：`https://github.com/ivy92888-glitch/sheet-keyweave.git`

## 2. 配置提交身份（仓库级，不改动全局设置）

```bash
cd /d/WorkBuddyDoc/sheet-keyweave

git config user.name  "你的显示名"
git config user.email "你的邮箱"
```

注意：提交邮箱会永久留在仓库历史里，请填写你愿意公开的邮箱。GitHub 账号设置中的 `Keep my email addresses private` 是常规隐私选项，需要的话可在账号设置里开启，并使用 GitHub 提供的 `noreply` 邮箱作为上面的 `user.email`。

## 3. 准备认证

HTTPS 推送需要凭据。若本机没有配置凭据管理器，推荐用 Personal Access Token：

1. GitHub → 头像 → `Settings` → `Developer settings` → `Personal access tokens` → `Tokens (classic)` → `Generate new token (classic)`
2. 勾选 `repo`，生成后**立刻复制**（只显示一次）
3. 推送时：用户名填 GitHub 账号名，密码栏粘贴 token（不要填账号密码）

Token 不要写进任何文件、不要提交到仓库。

## 4. 初始化并推送

```bash
cd /d/WorkBuddyDoc/sheet-keyweave

git init
git add .
git commit -m "Initial commit: generic VBA tabular cleaning toolkit"
git branch -M main
git remote add origin https://github.com/ivy92888-glitch/sheet-keyweave.git
git push -u origin main
```

Windows 下也可用 `cd /d D:\WorkBuddyDoc\sheet-keyweave`（CMD）进入目录，其余命令相同。

## 5. 推送前自检（建议每次都做）

```bash
git status
git ls-files
```

逐条核对 `git ls-files` 的输出，确认只包含下面这些文件，没有混入任何本地中间产物或原始素材：

```
.gitattributes
.gitignore
LICENSE
PUSH_GUIDE.md
README.md
docs/Settings.md
examples/sample.csv
src/DataCleaner.bas
```

原始素材、分析中间文件、含宏工作簿都保存在本仓库**之外的独立目录**，`.gitignore` 里也已做了防呆。若 `git ls-files` 里出现预期之外的文件，先 `git rm --cached <文件名>` 再提交。

## 6. 备选：不走命令行

如果 502 代理问题一时解决不了，可以用网页端上传：在新建的空仓库里点 `Add file → Upload files`，把本目录的 8 个文件按原有目录结构拖进去（`src/`、`docs/`、`examples/` 需保持层级），填写 commit message 后提交。

注意网页端上传**不会**自动包含以点开头的 `.gitignore` 和 `.gitattributes`，这两个文件需要在网页端用 `Add file → Create new file` 单独创建并粘贴内容。

## 7. 推送后验证

- 仓库首页能看到 README 渲染正常、LICENSE 被识别为 MIT
- `src/DataCleaner.bas` 可下载，且用记事本打开无乱码
- `Insights → Contributors` 里显示的身份是你期望的显示名
- 用 `git log --oneline` 确认只有一条初始提交，历史里没有多余文件

## 8. 502 排查方向

`CONNECT tunnel failed` 通常是本机代理或网络策略问题，不是 GitHub 账号问题。可依次尝试：换网络、检查系统代理设置、在 Git Bash 里执行 `git config --global --get http.proxy` 看是否配置了失效的代理、或换成网页端上传方案（第 6 步）。
