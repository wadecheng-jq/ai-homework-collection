# 🤖 AI作业收集系统 - 使用说明

> 用于收集全员AI作业，自动记录提交情况，并比对花名册生成未提交清单。

---

## 📦 文件结构

```
ai_homework_collection/
├── app.py                # 主服务程序（Flask网页服务）
├── check_submissions.py  # 核对脚本（比对花名册 → 未提交清单）
├── requirements.txt      # Python依赖列表
├── start.bat            # Windows一键启动脚本
├── templates/
│   ├── form.html        # 用户提交表单页面
│   └── admin.html      # 管理后台页面
├── static/
│   └── uploads/        # 附件上传存储目录
├── submissions.xlsx     # 自动生成的提交记录（首次运行自动创建）
└── 花名册.xlsx          # 人员花名册（放置于桌面，脚本自动读取）
```

---

## 🚀 快速启动

### 方式一：双击启动（推荐）

1. 双击运行 `start.bat`
2. 等待依赖安装完成（首次运行）
3. 浏览器访问：`http://localhost:5000`
4. 管理后台：`http://localhost:5000/admin`

### 方式二：命令行启动

```bash
cd C:\Users\wade\WorkBuddy\2026-06-05-19-46-24\ai_homework_collection
python app.py
```

---

## 📋 使用流程

### 第一步：分享提交链接

启动服务后，将以下链接发给全员：

```
http://[你的IP地址]:5000
```

> ⚠️ 注意：确保防火墙允许5000端口，同事才能访问。
> 查看本机IP：`ipconfig`（Windows）或 `ifconfig`（Mac/Linux）

### 第二步：监控提交进度

随时访问管理后台查看进度：

```
http://localhost:5000/admin
```

后台功能：
- 查看所有提交记录（实时刷新）
- 搜索提交人员
- 下载完整提交记录Excel
- 展开查看作业文字内容

### 第三步：截止后生成未提交清单

1. 截止时间（2026-06-08 18:00）过后，运行核对脚本：

```bash
cd C:\Users\wade\WorkBuddy\2026-06-05-19-46-24\ai_homework_collection
python check_submissions.py
```

2. 脚本自动：
   - 读取花名册（Desktop/花名册.xlsx）
   - 比对提交记录（submissions.xlsx）
   - 生成 `未提交人员清单_YYYYMMDD_HHMMSS.xlsx`
   - 生成 `未提交人员清单_YYYYMMDD_HHMMSS.txt`（可直接复制转发）

3. 打开Excel清单，按「未提交人员」逐一催办。

---

## ⚙️ 配置说明

### 修改截止时间

编辑 `app.py`，找到：

```python
DEADLINE = datetime.datetime(2026, 6, 8, 18, 0, 0)
```

修改日期时间即可，格式：`YYYY, M, D, H, M, S`

### 修改花名册路径

编辑 `check_submissions.py`，找到：

```python
ROSTER_PATH = os.path.join(BASE_DIR, "..", "..", "..", "Desktop", "花名册.xlsx")
```

修改为花名册文件的实际路径。

### 允许的文件类型

编辑 `app.py`，找到 `ALLOWED_EXTENSIONS`，可添加更多扩展名。

---

## 🔒 截止时间锁定机制

- **前端**：实时倒计时，截止后表单自动锁定（灰色遮罩 + 提示）
- **后端**：API层二次校验，即使绕过前端也无法提交
- **双重保障**：确保截止后无人能提交

---

## 🛠️ 常见问题

### Q：同事无法访问提交页面？
A：检查防火墙是否允许5000端口，或尝试关闭防火墙测试。
也可修改 `app.py` 最后一行，将 `port=5000` 改为其他端口。

### Q：提交的附件存在哪里？
A：存在 `static/uploads/` 目录下，以MD5重命名（原始文件名记录在Excel中）。

### Q：误删了 submissions.xlsx 怎么办？
A：停止服务后，将 `submissions.xlsx` 恢复至原目录，重新启动即可。

### Q：如何重启服务？
A：`Ctrl + C` 停止，再运行 `start.bat` 或 `python app.py`。

---

## 📊 交付物清单

| 交付物 | 状态 | 说明 |
|--------|------|------|
| 自定义网页表单 | ✅ 完成 | 姓名+圈子+作业内容+附件上传 |
| 截止时间自动锁定 | ✅ 完成 | 前后端双重锁定 |
| 提交数据后台（Excel） | ✅ 完成 | 自动记录，格式规范 |
| 管理后台页面 | ✅ 完成 | 实时查看、搜索、下载 |
| 自动核对脚本 | ✅ 完成 | 比对花名册→未提交清单 |
| 未提交清单（Excel） | ✅ 完成 | 含序号、姓名、备注列 |
| 未提交清单（文本） | ✅ 完成 | 可直接复制转发催办 |

---

*生成时间：2026-06-05 | WorkBuddy for Wade*
