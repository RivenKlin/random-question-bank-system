# 随机题库答题系统

一个基于 Flask 和 SQLite 的本地/服务器题库答题系统，支持 Excel 题库导入、随机练习、单选题、多选题、判断题、错题练习、收藏题练习、题目管理和答题结果检查。

## 功能

- Excel 自动导入单选题、多选题、判断题
- 按章节随机抽题
- 分别设置三种题型数量
- 左侧题号导航和答题进度恢复
- 错题练习和收藏题练习
- 管理端新增、编辑、删除题目
- SQLite 保存题库、收藏、错题、答题记录和服务器端答题进度
- 按设备标识与客户端 IP 保存未完成答题进度

## 本地运行

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python app.py
```

访问：

- 答题端：`http://127.0.0.1:5000/`
- 管理端：`http://127.0.0.1:5000/admin`

## 服务器运行

```bash
python3 -m venv .venv
. .venv/bin/activate
pip install -r requirements.txt
PORT=8044 waitress-serve --listen=0.0.0.0:8044 app:app
```

访问：`http://服务器IP:8044/`

生产环境建议使用 systemd、Nginx、HTTPS 和防火墙，仅开放需要的端口。管理员页面目前没有独立登录认证，不建议直接暴露到公网。

## 题库导入

将 `.xlsx` 文件放入 `已导入考试系统` 目录。工作表名称支持：

- `choice`：单选题
- `mchoice`：多选题
- `judgement`：判断题

重新导入可使用管理端的导入功能。
