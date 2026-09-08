# JSON 表格匹配复核工具

项目目录：`/Users/yyt/json_excel_review_tool`

## 启动

```bash
cd /Users/yyt/json_excel_review_tool
./start.sh
```

浏览器打开：<http://127.0.0.1:8765/>

## 停止

```bash
cd /Users/yyt/json_excel_review_tool
./stop.sh
```

## 重启

```bash
cd /Users/yyt/json_excel_review_tool
./restart.sh
```

默认服务只监听本机 `127.0.0.1`。如需局域网访问，可以这样启动：

```bash
BIND_HOST=0.0.0.0 PORT=8765 ./start.sh
```

服务日志位于 `work/server.log`，进程号位于 `work/server.pid`。上传文件、历史记录和导出文件分别保存在 `work/uploads`、`work/history` 和 `outputs`。
