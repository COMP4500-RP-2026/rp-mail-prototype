# RP Mail：macOS 试用说明

此版本是研究邮件流程原型。macOS 通过 SMTP 发信，不调用 Outlook 的 Windows 接口；不会自动读取回复、处理 survey 或更改 RP+。

## 从源代码运行

需要 Python 3.10+。在仓库目录运行：

```sh
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python desktop_app.py
```

公开邮箱查找依次尝试 Microsoft Edge、Google Chrome 和 Playwright Chromium。已安装 Edge 或 Chrome 时无需额外下载浏览器；都没有时运行：

```sh
.venv/bin/python -m playwright install chromium
```

## 配置发信账号

1. 在“Account settings”填写姓名、发件邮箱，并选择 SMTP。
2. 使用 Microsoft 365 时，向邮箱管理员确认 SMTP AUTH 是否允许，并取得获准使用的 public-client application ID 和 tenant ID。填入服务器、端口、登录邮箱、租户 ID 和应用 ID。默认 `smtp.office365.com:587`、STARTTLS 仅适用于相应 Microsoft 365 配置。
3. 点击“Check sending connection”。Microsoft 登录会显示一次性代码并打开登录网页；完成登录后，连接检查不会发送邮件。
4. 导入数据、核对收件人和推荐依据，先用“Test to myself”检查邮件，再处理获准联系的记录。

如果邮件服务支持密码式 SMTP，可选择该方式，并在启动程序前设置界面所填名称对应的环境变量。软件只保存环境变量名称，不保存密码。从 Finder 启动的 `.app` 通常不会继承 Terminal 中设置的变量。

## 打包与数据目录

在 macOS 上安装开发依赖并运行：

```sh
.venv/bin/python -m pip install -r requirements-dev.txt
.venv/bin/python -m PyInstaller --noconfirm RPMail.spec
```

产物是 `dist/RPMail.app`。macOS 上的工作清单、设置和发送历史保存在 `~/Library/Application Support/RPMail/`。升级时保留该目录，特别是 `send_history.sqlite3`。Windows 包必须在 Windows 上单独构建。

## 当前限制

- SMTP 不会在 Outlook 中创建草稿。发送成功提示表示服务端已接受，不等于研究人员已收到。
- 若学校禁止 SMTP AUTH，macOS 发信需要由管理员批准一个可用的发信方案；程序不会绕过机构策略。
- 自动邮箱查找只处理 RP+ 公开网页。找不到或无法读取的联系人须人工核对；不要将候选联系人直接视为关系确认人。
- 此仓库当前没有 survey submission 的自动导入功能，研究人员回复仍需在邮箱中查看。
