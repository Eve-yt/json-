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

## 违纪违法卷言证书证评测

左侧选择“违纪违法卷”，功能拆分为“黄金标准、数据获取、事件抽取评测、言证问题评测、书证问题评测、事件质量专项、人工匹配复核、综合报告”八个页面，并共用当前评测集和运行快照。侧边菜单会显示未准备、待评测、缺少输出、待复核和已完成等阶段状态。

- 创建、冻结和管理黄金评测集；
- 下载全中文的“违纪违法卷言证书证黄金标准标注模板.xlsx”，或导入、编辑事件/言证问题/书证问题标注；旧版英文标注类型仍可导入；
- 上传包含 `extract_transcript_events` 和 `issues` 的系统输出 JSON；
- 分别查看事件、言证和书证的严格/宽松精确率、召回率与 F1；
- 专项统计主体归属、字段准确率、事件完整率、摘要合格率、无关信息率、重复率、拆分率、合并准确率、书证关联准确率和输入范围限制；
- 在独立页面人工确认或驳回自动匹配，并导出综合 Excel 评测报告。

评测数据保存在 `work/review_tool.db`，报告输出到 `outputs`。

## 主体身份卷评测

左侧选择“主体身份卷”，功能拆分为“黄金标准、数据获取、文件分类评测、职位抽取评测、任职对齐评测、身份材料评测、业务规则评测、综合报告”八个页面，并共用当前评测集和运行快照。

- 文件分类精确率、召回率、F1和文件类型准确率；
- 职位、职级、分工抽取精确率、召回率、F1及字段准确率；
- 任职经历分组准确率、错误拆分率和错误合并率；
- 特殊身份、其他身份识别F1及关联材料准确率；
- 业务规则通过、失败和未执行数量。

黄金标准模板包含“案件信息、文件分类、职位抽取、任职分组、身份材料、业务规则”六个工作表，文件名、列名、示例和选项均使用中文；导入时仍兼容旧版英文列名。系统输出可以通过 JSON 上传或从主体身份接口拉取，历史运行和评测报告保存在本地数据库及 `outputs` 目录。

运行专项测试：

```bash
/Users/yyt/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 work/test_evidence_evaluation.py
```
