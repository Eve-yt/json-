# 案件审查系统 Agent 全流程评测平台搭建方案

编写日期：2026-10-19。依据本地交付包静态代码阅读；未启动服务、连接数据库或调用模型。

## 1. 目标与结论

建设覆盖“输入材料→提示词→模型决策→工具执行→规则判断→结果持久化→Java 展示”的评测平台。核心目标是解释：结果错在哪里、哪一步首次出现可观测错误、修改哪个组件能改善，以及改善是否引入回归。

推荐结构：Python 业务侧统一采集 SDK + 独立评测服务和数据存储 + 现有 JSON/Excel 复核工具的黄金标准与人工裁决能力。沿用已有规则 playground 作为试点入口，逐步统一其执行器与生产执行器。Java 侧负责入口、回调、结果同步和人工确认的端到端关联。

第一阶段先采集和离线评分，后续再加入暂停、替换输入、重试等主动干预。运行干预会改变被测流程，必须独立记录策略版本，不能与原始基线实验混合。

## 2. 阅读范围和证据边界

- JProject4SL：可见 Python 源码，重点阅读 app 接口、handler、procedure 编排、Agent 工具层、LLM 客户端、主体身份流水线、证据流水线、playground 和相关测试。
- intelli-judge-service：当前目录为部署包，主要包含 Spring Boot JAR、前端 dist、配置与 SQL，没有完整 Java/Vue 工程源码。检查了 JAR 类目录，并用 javap 阅读关键类签名、字节码与接口常量。因此可确认组件与部分调用关系，不能等同完整源码审查。
- json_excel_review_tool：复用当前仓库的黄金标准、事件/身份评分、人工匹配复核和 Excel 导出。
- 交接文档只作背景材料，以当前可执行代码为准。例如程序卷材料缺失处理已采用统一分流，不能直接沿用交接文档中的旧容缺材料描述。
- 未验证部署地址、模型中台、数据库现状、接口互通和实际质量分数；本方案不声称完成这些验证。

## 3. 当前实际执行链

### 3.1 Java 业务入口

JAR 中可见 CapacityProcessController、CapacityProcessCallbackController、ReviewQuestionClient、RuleConfigProxyService、CaseReviewQuestionServiceImpl，以及人工确认请求类和 CaseReviewQuestionSnapshot。

字节码中可见 /review/procedure、/subject_identity/review、/violation/review、/review/procedure/result、/procedure/rule/result、/review/questions/sync、/cpc/callback 等路径。不同组件使用的前缀存在差异，接入时需确认网关和代理重写，不能把字符串常量直接当成已经验证的线上 URL。

评测应覆盖：用户提交→任务流转→调用 Python→异步回调→查询/同步结果→展示→人工确认。原始 AI 输出、Java 同步结果和人工修改后的结果应分别保存，防止拿修正后的数据评估模型。

### 3.2 程序卷

入口 src/app/api/business/procedure.py，经 handler 进行分类与业务执行。核心在 src/business/material_review/procedure/service.py：

1. 加载规则并按业务分区过滤。
2. should_run_rule 判断案件条件。
3. 读取材料并构造执行队列。
4. 材料类型检测规则始终执行；缺少任意来源材料的普通规则静默跳过；必要时构造合成类型检测规则。
5. 规则并发运行，进入 pre_check→planner→worker→solver。
6. Worker 可走快速通道，或进入 ReAct 循环和视觉工具。
7. Solver 含解析、跨实例比较及规则防护逻辑。
8. 结果落库，API 再聚合、过滤和截断部分描述。

重要事实：规则计算异常、落库异常和最终展示异常是不同错误；当前部分落库异常被隔离后任务仍继续，不能以“任务结束”代替“结果完整”。

### 3.3 主体身份卷

src/app/service/handlers/review_subject_identity.py 明确分为 file_classify、position_extract、position_align、identity_search 四个阶段。最后阶段还执行独立主体身份规则，规则服务复用 run_pipeline。

因此需同时覆盖固定流水线与规则 Agent，不能只给 ReAct 加日志。

### 3.4 言证书证

review_transcript/pipeline.py 与 stages 下模块组织页面文本、问答切分、事件抽取、事件关联、书证检索和问题审查。ModelGateway 管理模型调用、校验、修复等行为；llm_compat 最终转入公共 LLM 客户端。

已有 infrastructure/observability.py 提供可选诊断事件，但主要是无正文的阶段、耗时、状态数据，观察者为进程级变量。适合接成统一采集适配器，不能直接假定它具有每次实验隔离及完整输入输出能力。

## 4. 已有能力与必须补齐的缺口

| 能力 | 当前基础 | 需要补齐 |
|---|---|---|
| 规则工作台 | procedure/playground.py、/playground/api/* | 批量实验、黄金标准、版本对比、生产执行一致性 |
| 模型调用记录 | chat.py 的 recorder、context_meta 的 ContextVar | 请求开始、失败、取消、每次重试、原始 tool_calls |
| 工具证据 | ReactLoop 的 Evidence、BaseTool 参数校验 | 原始参数、系统注入参数、最终参数分别记录；快速通道覆盖 |
| 规则标识 | caseId/bizId/ruleId/phase 等 metadata | experiment/run/span/parent_span 标识，合成规则来源 |
| 提示词 | 代码常量、Jinja、PROMPT_TEMPLE、规则 INPUT/PLAN | 渲染前模板和最终请求快照、哈希、来源及版本 |
| 业务评分 | 当前复核工具的事件和身份指标 | 过程标注、规则激活评分、端到端同步一致性 |

发现的具体缺口：

- playground._run_rule_task 在可能生成 Planner 之后才设置 recorder 和规则 metadata，首次 Planner 请求可能不在本次记录中；应在任何模型调用之前建立上下文。
- chat._record_llm_call 仅在成功解析响应后触发，失败尝试不会形成完整调用记录；当前 rec 未传 raw response，纯 tool_calls 响应可能丢失关键动作信息。
- Playground 自编排 plan→workers→solve；生产还有条件过滤、材料执行队列、pre_check 等前置分支。单测成功不能直接证明生产路径成功。
- 当前运行 JSON 主要在结束时保存，运行中只返回 running；进程异常后可能缺少中间记录。需要增量写事件和中断恢复状态。
- 图片在 playground 中被替换成大小占位符，便于展示，但不能复现。必须另外保存受控的图片/页面产物引用与内容哈希。
- trace_id 有时使用 caseId:bizId，有时使用 runId。业务键不应替代独立运行标识。
- 规则可能按同一个规则 ID 生成合成材料类型检测结果，评测键必须加入 rule_kind、来源材料和实例，不能只以 ruleId 计数。

## 5. 目标架构

```text
Java 提交/回调/人工确认 ──────── 端到端业务关联
              │
Python handler / scheduler
              │
   规则筛选 / 固定流水线 / Agent 编排
              │
   Prompt → LLM → Tool / FastPath → Solver → 持久化
              │
        统一采集 SDK（追加事件）
              │
      本地缓冲 / 批量投递 / 去重
              │
        独立评测服务与存储
      ┌───────┼────────┐
  实验与版本  分步评分/回归  链路查看/人工复核
              │
      复用现有黄金标准和报告引擎
```

采集失败不得改变业务结论。实验环境可启用严格模式：采集不全则实验标记为 incomplete，不能显示成通过。生产环境采用受控采样、内容访问权限和明确保留期限。

建议先沿用现有 Python/FastAPI 技术体系，平台使用独立数据库或 schema，正文和 PDF/图片产物采用独立文件存储。单机试点可用 SQLite 与本地目录；团队部署再按并发需求使用支持多进程写入的数据库与独立任务 Worker。此处是架构建议，不是已经验证的部署配置。

## 6. 统一采集模型

层级：Experiment→CaseRun→Stage→Rule/MaterialInstance→LLMAttempt/ToolCall→Artifact。

每条事件至少包含：event_id、schema_version、experiment_id、run_id、span_id、parent_span_id、event_type、service、case_id、biz_id、rule_id、rule_kind、material_instance_id、file_id、page_range、attempt、sequence、时间、状态、错误类型、输入/输出产物引用。

- event_id 用于重复投递去重；同一 span 允许 started/finished/error 等事件。
- span 关系表达嵌套；多输入合并步骤另加 links 表达证据依赖，不能只靠时间顺序。
- 每次真实模型重试独立计时计费，逻辑调用和重试尝试分层。
- 标记 route=react/sft_fast_path/material_type_precheck 等路径，以及路由原因。
- 记录 queue_ms、execution_ms、parse_ms、persist_ms，区分排队与计算。
- 状态区分 success、error、cancelled、skipped、incomplete；跳过必须携带原因。
- async task、调度队列、线程和跨服务都显式传递 run/span 标识；线程池需逐处验证上下文传播，不能只依赖 HTTP 中间件。

模型记录：最终序列化 messages、可用工具 schema、模型标识与服务版本、采样参数、响应 content/tool_calls/finish_reason/usage、异常和响应产物引用。记录可见决策和证据，不要求模型暴露内部思维链。

工具记录：LLM 提出的原始参数、INJECTED_ARGS 覆盖项、校验后的最终参数、工具输出、对应材料和页码、耗时和异常。否则无法区分模型参数错误与系统注入错误。

提示词记录：template_key、source=code/jinja/db/rule、模板内容哈希、模板快照、变量输入、最终消息快照。复现需要实际内容，只有哈希不够。

## 7. 具体改造接入点

| 位置 | 改造内容 |
|---|---|
| app/service/handlers、schedule | 建立实验上下文，跨异步调度传递，记录取消与任务终态 |
| llm/context_meta.py | 在已有 metadata 上扩展 run/span；增加可恢复的上下文作用域 |
| llm/client/chat.py、infra/retry.py | 统一模型调用与 attempt 事件；记录失败、原始 tool_calls、排队和使用量 |
| procedure/service.py | 规则条件、材料存在性、静默跳过、合成规则、结果持久化分别发事件 |
| pipeline/orchestrator.py | pre_check/planner/workers/solver 阶段 span，共用生产执行路径 |
| pipeline/planner.py | 记录 cache hit、现场生成、兜底 instruction、最终 plan 和缓存来源 |
| agentkit/agent/react/loop.py | 每轮可用动作、模型动作、JSON 参数错误、迭代上限和强制摘要 |
| agentkit/tools/base.py | 参数校验、执行结果和异常；直接快速调用另补适配 |
| pipeline/worker.py、fast_path.py | 记录路由理由、实际图像/页面覆盖、分块完整性、SFT 工具输出 |
| pipeline/solver.py、solver_guards.py | 原始判定、解析、guard 修正前后，以及最终结构化结果 |
| procedure/playground.py | 最先挂载 recorder；逐步替换自编排为共享执行器；事件增量保存 |
| subject_identity handler/progress | 四阶段产物快照、分组依赖、规则激活与独立规则输出 |
| review_transcript ModelGateway/observability | 桥接诊断事件，补正文引用、校验修复、截断重分片及检索候选 |
| Java 入口/回调/同步/确认组件 | 跨服务关联、输入输出快照、同步差异、人工修订与原始结果分离 |

Java 改造需要完整工程源码；在源码到位前，先用测试请求关联表和 Python 侧采集完成试点，Java 部分仅做接口级验证，明确覆盖范围。

## 8. 评测体系

### 8.1 规则激活与过程正确性

建立“案件×规则×材料实例”的期望执行矩阵：应执行、条件不满足、材料缺失、业务不适用。分别测应执行召回率、错误执行率、跳过原因准确率。

重要：不存在结果行不等于通过；缺材料而跳过不等于漏执行；已完成任务不等于全部结果已落库。分母来自冻结的黄金执行矩阵，不能只来自实际产生的结果。

Plan 不做逐字匹配，评估约束覆盖、来源材料、抽取目标、single_file/cross_file 模式是否满足业务要求。工具步骤评估允许动作集合、关键参数、输入材料范围和证据完整性，不要求只有一条固定调用顺序。

模型调用分开衡量：传输成功、结构合法、业务正确。JSON 修复成功不能替代内容正确。

### 8.2 工具与视觉能力

签名/印章/捺印等评估检出与误报、字段归属、文件实例、页码范围及不完整分块率。日期工具增加日期归一化准确率；需要位置框时另补框级标注，不能用页级标签冒充检测精度。

检索同时记录候选集和最后选中集合，计算有标注条件下的 Recall@K、关联精确率及原文引用支持率。OCR 失败后视觉兜底、截断后重分块都保留为显式路径。

### 8.3 主体身份与言证书证

主体身份沿用分类、职位/职级/分工字段、任职分组、身份材料关联指标，并追加阶段输入输出对比。分组准确性按成员关系评估，不要求 group_id 字符串一致。

言证书证沿用事件、问题、材料关联的严格/宽松 P/R/F1，追加抽取范围、事件主体、事实一致性、重复/拆分/合并、检索覆盖与证据支持。记录窗口、事件、组、材料之间的稳定标识映射。

### 8.4 系统层指标

分别统计计算成功率、结果落库成功率、Java 同步一致率、回调成功率、端到端可展示率；统计 p50/p95 延迟、真实 token、重试率、缓存命中率和重复调用。

按案件平均与按样本汇总同时呈现，按材料类型/规则类型/难度分层。不建议一开始压成单一总分。

## 9. 黄金数据、实验与错误定位

黄金标准分层：案件信息→材料分类/页范围→预期规则激活→中间产物→预期业务结论与依据。允许只标注某些层级，缺少标注显示未评测，不能显示通过。

版本冻结包含：PDF 和 OCR 产物、分类结果、案件上下文、规则及 PLAN 快照、提示词、模型/适配器配置、代码标识、工具 schema、路由开关、缓存模式、评分器版本。没有 Git 元信息的部署包使用构建标识和文件哈希。

公平比较：同一冻结评测集，分别控制代码/Prompt/模型/规则变量；缓存至少区分冷启动与固定快照命中。保留真实模型多次重复运行的分布，对稀有规则报告样本数。

错误定位展示“最早可观测偏差”与依赖路径，不能仅凭时间先后宣称因果。例如材料分类漏检→规则未激活→最终漏报是待验证归因，可通过固定正确分类后重跑规则进行验证。

AI 裁判主要用于摘要和语义等难以确定性评分的任务。保存裁判模型、rubric、输入、输出和证据；用人工已裁决样本校准，低置信度/争议样本进入人工复核。

## 10. 重放与干预

提供四种模式：

1. 历史查看：只读历史链路和产物，不调用模型。
2. 模型响应回放：匹配请求内容、模型参数、工具 schema 和上下文后返回保存响应，用于工程回归，不代表真实模型质量。
3. 节点重跑：固定上游快照，只重跑 Planner/Tool/Solver/事件阶段，创建新的 child run，便于隔离变量。
4. 全链路实验：从原始材料开始执行，测真实业务结果。

节点干预可支持 Prompt/Plan 草稿覆盖、mock 工具输出、替换材料快照、指定规则、人工暂停。所有覆盖只在实验上下文生效；发布到正式规则表另走已有发布权限及版本流程。现有 playground 的 save-plan 不应成为实验保存操作的隐含副作用。

环境隔离采用独立测试库/schema 和产物空间；pg_ 前缀可保留用于识别，但不能替代隔离。每次实验分配独立 bizId 并保存源案件映射，因为当前业务存在按 caseId/bizId 清理重写结果的路径。回调指向实验接收器，避免驱动正式案件状态。

## 11. 平台页面和数据表

页面：实验总览、运行链路、提示词/Plan 调试台、工具与规则专项、黄金标准/人工复核、A/B 版本对比、错误回归集。

链路详情采用左侧执行树、中央输入输出差异、右侧原文与黄金标准。支持按错误类型、规则、材料实例、模型、路由筛选；并行步骤显示为树或泳道，不强行排成唯一执行顺序。

核心数据表建议：

- eval_dataset / eval_dataset_version / eval_sample：冻结样本与标注。
- eval_component_version：代码、Prompt、规则、模型和工具快照索引。
- eval_experiment / eval_run：实验参数、案件运行、运行模式与父运行。
- eval_span / eval_event / eval_span_link：执行结构、追加事件与依赖。
- eval_artifact：正文、PDF、图像、OCR 和中间结果引用及哈希。
- eval_score / eval_review：自动评分、裁判版本和人工裁决。
- eval_issue / eval_regression_case：错误分类和回归样本。

原始事件追加保存，评分和人工结论可版本化更新；不能覆盖原始输出。对采集丢失显示 completeness 状态。数据库事务或本地持久缓冲保障事件投递，业务采集与评测服务解耦。

## 12. 落地顺序与验收

### 阶段一：程序卷单规则完整链路

复用 playground，补齐最早上下文、LLM 失败与重试、tool_calls、参数注入、快速通道、规则条件和持久化事件。提供历史链路和简单黄金判定。

验收：冷 Planner、缓存命中、ReAct、快速通道、条件跳过、材料缺失、非法工具参数、模型重试、取消及落库失败均有明确事件；并行规则不串链；采集器故障不改变业务返回；运行进程中断后记录标记为中断或不完整。

### 阶段二：生产一致执行与批量对比

共享生产执行器，建立冻结评测集、唯一运行键、Prompt/Plan/规则快照、批量实验、版本差异和错误回归。将现有 Excel 评分与复核逻辑抽成可调用评分器，保留原始分数与复核后分数。

验收：相同数据与配置下，工作台和生产入口的执行分流及结果具有可解释的一致性；历史实验不随数据库规则更新而改变；两版差异可下钻到案件、规则与材料实例。

### 阶段三：主体身份与言证书证

接入身份四阶段及独立规则、证据 ModelGateway/observer、检索候选与事件关联，补各自过程黄金标准和回放适配器。

验收：可分别验证抽取、分组、检索和最终问题，不只查看总 F1；节点重跑创建新运行且不污染上游快照。

### 阶段四：Java 端到端和发布门槛

取得 Java 工程源码后接入口、回调、同步和人工确认；定义关键错误不得回归、业务指标不低于基线及耗时预算等发布门槛。具体阈值由冻结基线与业务成本确定，不预设未经测量的百分比。

验收：可区分 Python 正确但 Java 未同步、同步正确但人工修改、模型错误但展示被兜底等情况；发布报告包含质量、过程、完整性及性能。

## 13. 首批样本建议

先选小规模、人工可完全复核的代表集，覆盖：单材料多实例、跨材料日期比较、签名/印章视觉快速通道、规则条件不满足、材料缺失触发合成类型检测、OCR 失败兜底、Plan 缓存旧数据、模型 JSON 修复、工具参数错误、重试与取消、落库失败、Java 同步差异。

第一交付物应是：一个案件的一条规则，从输入与规则适用性，到完整 Prompt、模型动作、工具执行、Solver 修正、数据库结果和黄金判定均可查看；随后再扩展批量和其他业务。

## 14. 关键源码索引

Python 根目录：/Users/yyt/shenli-yt/JProject4SL/

- src/llm/client/chat.py：公共模型请求与成功调用 recorder。
- src/llm/context_meta.py：业务 metadata 与协程级 recorder。
- src/agentkit/agent/react/loop.py：工具选择、参数注入、证据与结束逻辑。
- src/agentkit/tools/base.py：工具 schema、参数校验、执行异常。
- src/business/material_review/procedure/service.py：当前有效实现从后半部分开始，前半部分大量为注释旧代码。
- src/business/material_review/procedure/pipeline/{orchestrator,planner,worker,fast_path,solver,solver_guards}.py。
- src/business/material_review/procedure/playground.py：单规则工作台、自编排、运行 JSON、Plan 覆盖和写回。
- src/app/api/business/playground.py、src/app/main.py：工作台路由与条件注册。
- src/business/material_review/prompt_store.py：数据库 Prompt 及代码侧 Prompt 的边界。
- src/app/service/handlers/review_subject_identity.py：身份四阶段。
- src/business/material_review/subject_identity/rules/service.py：独立身份规则。
- src/business/material_review/review_transcript/pipeline.py、infrastructure/{llm,observability}.py：证据流水线与诊断。
- tests/test_procedure_routing.py：SFT/solver 路由与视觉分块的既有边界案例。

Java 包：/Users/yyt/shenli-yt/intelli-judge-service/lib/intelli-judge-service-2.0.0.jar。

复核平台根目录：/Users/yyt/json_excel_review_tool/，重点复用 work/evidence_evaluation.py、work/subject_identity_evaluation.py 与 work/matching_review_tool.py。
