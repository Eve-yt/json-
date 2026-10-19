"""Offline evaluation for violation-volume verbal and documentary evidence.

The module is intentionally independent from the procedure-volume matcher.  It
stores gold annotations and immutable run snapshots in the application's SQLite
database, then compares the two public result channels returned by the review
service: ``extract_transcript_events`` and ``issues``.
"""

import base64
import json
import re
import sqlite3
import time
import uuid
from collections import defaultdict
from difflib import SequenceMatcher
from io import BytesIO
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


GOLD_COLUMNS = [
    ("kind", "标注类型", "事件 / 言证问题 / 书证问题；兼容旧值 event / verbal_issue / documentary_issue"),
    ("case_id", "案件编号", "案件唯一编号"),
    ("file_id", "文件ID", "系统返回的 file_id"),
    ("item_id", "材料项ID", "系统返回的 item_id"),
    ("item_name", "材料名称", "笔录或书证名称"),
    ("page_no", "页码", "证据所在页码"),
    ("source_text", "原始文本", "问答或书证原文"),
    ("subject", "主体", "被调查人、证人或其他主体"),
    ("behavior", "行为/问题类型", "兼职取酬、收受礼金等"),
    ("event_time", "时间", "事件时间"),
    ("amount", "金额", "金额或财物价值"),
    ("location", "地点", "事件地点"),
    ("related_object", "关联对象", "相关人员、单位或财物"),
    ("summary", "标准事件或问题", "人工确认的标准描述"),
    ("evidence_text", "证据原文", "支持结论的原文片段"),
    ("event_cluster_id", "事件簇", "同一事实使用相同编号"),
    ("merge_relation", "合并关系", "独立/应合并/不可合并/重复/嵌套"),
    ("expected_conclusion", "预期结论", "问题应有的判断结论"),
    ("issue_scene", "问题场景", "兼职取酬、收受礼金、金额矛盾等"),
    ("document_status", "书证存在状态", "存在、缺失、无法判断"),
    ("related_files", "关联书证", "多个文件用 | 分隔"),
    ("input_limited", "输入范围限制", "是/否；如关键信息不在书证前两页"),
    ("limitation_reason", "匹配限制原因", "前两页限制、未读取材料等"),
    ("irrelevant_info", "无关信息", "页码、被谈话人等不应进入结果的信息"),
    ("split", "数据分组", "开发集/回归集/盲测集"),
    ("annotation_status", "标注状态", "初标/复核/仲裁/已确认"),
    ("notes", "备注", "标注或仲裁说明"),
]


def _text(value):
    if value is None:
        return ""
    if isinstance(value, (dict, list)):
        return json.dumps(value, ensure_ascii=False, sort_keys=True)
    return str(value).strip()


def _norm(value):
    value = _text(value).lower()
    value = value.replace("（", "(").replace("）", ")")
    return re.sub(r"[\s\W_]+", "", value, flags=re.UNICODE)


def _similarity(left, right):
    left, right = _norm(left), _norm(right)
    if not left or not right:
        return 0.0
    if left in right or right in left:
        return min(len(left), len(right)) / max(len(left), len(right)) * 0.35 + 0.65
    return SequenceMatcher(None, left, right).ratio()


def _bool(value):
    return _text(value).lower() in {"1", "true", "yes", "是", "受限"}


KIND_ALIASES = {
    "event": "event", "事件": "event",
    "verbal_issue": "verbal_issue", "言证问题": "verbal_issue",
    "documentary_issue": "documentary_issue", "书证问题": "documentary_issue",
}


def _unwrap_result(payload):
    current = payload
    for _ in range(3):
        if not isinstance(current, dict) or "data" not in current:
            break
        current = current.get("data")
    return current if isinstance(current, dict) else {}


def _event_text(event):
    if not isinstance(event, dict):
        return _text(event)
    preferred = [
        event.get(key) for key in (
            "summary", "description", "event", "content", "event_content",
            "事件", "事件描述", "事实描述",
        ) if event.get(key) not in (None, "")
    ]
    return " ".join(_text(value) for value in preferred) or _text(event)


def _pick(item, *names):
    if not isinstance(item, dict):
        return ""
    for name in names:
        if item.get(name) not in (None, ""):
            return _text(item.get(name))
    return ""


def flatten_model_result(payload):
    """Return normalized model events/issues while tolerating response variants."""
    data = _unwrap_result(payload)
    events = []
    event_groups = data.get("extract_transcript_events") or data.get("extractTranscriptEvents") or []
    if isinstance(event_groups, dict):
        event_groups = [event_groups]
    for group in event_groups if isinstance(event_groups, list) else []:
        if not isinstance(group, dict):
            group = {"events": [group]}
        group_events = group.get("events") or group.get("event_list") or []
        if isinstance(group_events, (str, dict)):
            group_events = [group_events]
        for event in group_events if isinstance(group_events, list) else []:
            event_dict = event if isinstance(event, dict) else {"summary": _text(event)}
            events.append({
                "id": uuid.uuid4().hex,
                "kind": "event",
                "fileId": _pick(group, "file_id", "fileId"),
                "itemId": _pick(group, "item_id", "itemId"),
                "itemName": _pick(group, "item_name", "itemName", "file_name", "fileName"),
                "subject": _pick(event_dict, "subject", "主体", "person", "actor"),
                "behavior": _pick(event_dict, "behavior", "行为", "event_type", "eventType", "type"),
                "eventTime": _pick(event_dict, "time", "event_time", "eventTime", "时间"),
                "amount": _pick(event_dict, "amount", "金额"),
                "location": _pick(event_dict, "location", "地点"),
                "relatedObject": _pick(event_dict, "related_object", "relatedObject", "对象", "object"),
                "summary": _event_text(event_dict),
                "raw": event,
            })
    issues = []
    raw_issues = data.get("issues") or []
    if isinstance(raw_issues, dict):
        raw_issues = [raw_issues]
    for issue in raw_issues if isinstance(raw_issues, list) else []:
        if not isinstance(issue, dict):
            issue = {"description": _text(issue)}
        issue_type = _pick(issue, "type", "issue_type", "issueType").upper()
        if issue_type == "VERBAL_EVIDENCE":
            kind = "verbal_issue"
        elif issue_type == "DOCUMENTARY_EVIDENCE":
            kind = "documentary_issue"
        else:
            continue
        related = issue.get("related_files") or issue.get("relatedFiles") or []
        if isinstance(related, str):
            related = [related]
        issues.append({
            "id": uuid.uuid4().hex,
            "kind": kind,
            "type": issue_type,
            "summary": _pick(issue, "description", "summary", "content"),
            "subject": _pick(issue, "subject", "主体", "person"),
            "conclusion": _pick(issue, "conclusion", "result", "结论"),
            "documentStatus": _pick(issue, "document_status", "documentStatus", "evidence_status", "evidenceStatus", "书证状态"),
            "relatedFiles": [_text(value) for value in related if _text(value)],
            "raw": issue,
        })
    return {"events": events, "issues": issues, "all": events + issues}


def _field_score(gold, model):
    summary_score = max(
        _similarity(gold.get("summary"), model.get("summary")),
        _similarity(gold.get("evidenceText"), model.get("summary")),
    )
    if gold["kind"] == "event":
        weights = {
            "subject": 0.20, "behavior": 0.20, "eventTime": 0.10,
            "amount": 0.10, "location": 0.05, "relatedObject": 0.05,
        }
        score = summary_score * 0.30
        available = 0.30
        for field, weight in weights.items():
            if _norm(gold.get(field)):
                score += _similarity(gold.get(field), model.get(field)) * weight
                available += weight
        return score / available if available else summary_score
    related_score = _similarity("|".join(gold.get("relatedFiles") or []), "|".join(model.get("relatedFiles") or []))
    subject_score = _similarity(gold.get("subject"), model.get("subject"))
    available = 0.75
    score = summary_score * 0.75
    if _norm(gold.get("subject")):
        score += subject_score * 0.15
        available += 0.15
    if gold["kind"] == "documentary_issue" and gold.get("relatedFiles"):
        score += related_score * 0.10
        available += 0.10
    return score / available


def _strict_match(gold, model):
    fields = ["subject", "behavior", "eventTime", "amount", "location", "relatedObject"] if gold["kind"] == "event" else ["subject"]
    for field in fields:
        expected = _norm(gold.get(field))
        if expected and _similarity(expected, model.get(field)) < 0.90:
            return False
    return _similarity(gold.get("summary") or gold.get("evidenceText"), model.get("summary")) >= 0.82


def evaluate(gold_rows, payload, loose_threshold=0.58):
    parsed = flatten_model_result(payload)
    models = parsed["all"]
    candidates = []
    for gi, gold in enumerate(gold_rows):
        if gold.get("inputLimited"):
            continue
        for mi, model in enumerate(models):
            if gold["kind"] != model["kind"]:
                continue
            score = _field_score(gold, model)
            if score >= loose_threshold:
                candidates.append((score, gi, mi))
    candidates.sort(reverse=True)
    used_gold, used_model, matches = set(), set(), []
    for score, gi, mi in candidates:
        if gi in used_gold or mi in used_model:
            continue
        used_gold.add(gi)
        used_model.add(mi)
        matches.append({
            "goldIndex": gi, "modelIndex": mi, "score": round(score, 4),
            "strict": _strict_match(gold_rows[gi], models[mi]),
            "reviewRequired": score < 0.78,
        })

    details = []
    matched_by_gold = {item["goldIndex"]: item for item in matches}
    matched_by_model = {item["modelIndex"]: item for item in matches}
    for index, gold in enumerate(gold_rows):
        match = matched_by_gold.get(index)
        if gold.get("inputLimited"):
            status = "input_limited"
        else:
            status = "matched" if match else "missed"
        details.append({"side": "gold", "status": status, "gold": gold, "model": models[match["modelIndex"]] if match else None, **(match or {})})
    for index, model in enumerate(models):
        if index not in matched_by_model:
            details.append({"side": "model", "status": "false_positive", "gold": None, "model": model})

    kind_metrics = {}
    for kind in ("event", "verbal_issue", "documentary_issue"):
        gold_indices = [i for i, row in enumerate(gold_rows) if row["kind"] == kind and not row.get("inputLimited")]
        model_indices = [i for i, row in enumerate(models) if row["kind"] == kind]
        kind_matches = [row for row in matches if row["goldIndex"] in gold_indices]
        strict = sum(1 for row in kind_matches if row["strict"])
        tp, gold_total, model_total = len(kind_matches), len(gold_indices), len(model_indices)
        precision = tp / model_total if model_total else (1.0 if not gold_total else 0.0)
        recall = tp / gold_total if gold_total else 1.0
        f1 = 2 * precision * recall / (precision + recall) if precision + recall else 0.0
        kind_metrics[kind] = {
            "goldTotal": gold_total, "modelTotal": model_total, "matched": tp,
            "strictMatched": strict, "precision": precision, "recall": recall, "f1": f1,
        }

    event_gold_indices = [i for i, row in enumerate(gold_rows) if row["kind"] == "event" and not row.get("inputLimited")]
    event_matches = [row for row in matches if row["goldIndex"] in event_gold_indices]
    subject_expected = [row for row in event_matches if _norm(gold_rows[row["goldIndex"]].get("subject"))]
    subject_correct = sum(
        _similarity(gold_rows[row["goldIndex"]].get("subject"), models[row["modelIndex"]].get("subject")) >= 0.90
        for row in subject_expected
    )
    duplicate_models = 0
    split_gold = set()
    for score, gi, mi in candidates:
        if mi in used_model and matched_by_model.get(mi, {}).get("goldIndex") != gi:
            continue
        if gi in used_gold and matched_by_gold.get(gi, {}).get("modelIndex") != mi and score >= 0.72:
            duplicate_models += 1
            split_gold.add(gi)
    event_model_total = len(parsed["events"])
    event_field_map = {"subject": "subject", "behavior": "behavior", "eventTime": "eventTime", "amount": "amount", "location": "location", "relatedObject": "relatedObject"}
    event_fields = {}
    for label, field in event_field_map.items():
        applicable = [row for row in event_matches if _norm(gold_rows[row["goldIndex"]].get(field))]
        correct = sum(_similarity(gold_rows[row["goldIndex"]].get(field), models[row["modelIndex"]].get(field)) >= .90 for row in applicable)
        event_fields[label] = {"correct": correct, "total": len(applicable), "accuracy": correct / len(applicable) if applicable else 1.0}
    complete_count = 0
    summary_qualified = 0
    noise_count = 0
    for row in event_matches:
        gold, model = gold_rows[row["goldIndex"]], models[row["modelIndex"]]
        required = [field for field in event_field_map.values() if _norm(gold.get(field))]
        complete_count += all(_norm(model.get(field)) for field in required)
        gold_summary, model_summary = _text(gold.get("summary")), _text(model.get("summary"))
        summary_qualified += _similarity(gold_summary, model_summary) >= .60 and len(model_summary) <= max(40, len(gold_summary) * 3)
        noise_tokens = [part.strip() for part in re.split(r"[|,，;；]", _text(gold.get("irrelevantInfo"))) if part.strip()]
        has_noise = any(token in model_summary for token in noise_tokens) or bool(re.search(r"第?\s*\d+\s*页|被谈话人[:：]", model_summary))
        noise_count += has_noise
    cluster_groups = defaultdict(list)
    for gi in event_gold_indices:
        cluster = _text(gold_rows[gi].get("eventClusterId"))
        if cluster: cluster_groups[cluster].append(gi)
    merge_groups = [indices for indices in cluster_groups.values() if len(indices) > 1 or any(_text(gold_rows[i].get("mergeRelation")) == "应合并" for i in indices)]
    merge_correct = 0
    for indices in merge_groups:
        related_models = {mi for score, gi, mi in candidates if gi in indices and score >= loose_threshold}
        merge_correct += len(related_models) == 1
    verbal_matches = [row for row in matches if gold_rows[row["goldIndex"]]["kind"] == "verbal_issue"]
    documentary_matches = [row for row in matches if gold_rows[row["goldIndex"]]["kind"] == "documentary_issue"]
    verbal_description_accuracy = sum(_similarity(gold_rows[row["goldIndex"]].get("summary"), models[row["modelIndex"]].get("summary")) >= .82 for row in verbal_matches) / len(verbal_matches) if verbal_matches else 1.0
    documentary_relation_accuracy = sum(_similarity("|".join(gold_rows[row["goldIndex"]].get("relatedFiles") or []), "|".join(models[row["modelIndex"]].get("relatedFiles") or [])) >= .75 for row in documentary_matches if gold_rows[row["goldIndex"]].get("relatedFiles"))
    documentary_relation_total = sum(bool(gold_rows[row["goldIndex"]].get("relatedFiles")) for row in documentary_matches)
    limited = sum(1 for row in gold_rows if row.get("inputLimited"))
    summary = {
        "event": kind_metrics["event"],
        "verbalIssue": kind_metrics["verbal_issue"],
        "documentaryIssue": kind_metrics["documentary_issue"],
        "subjectAccuracy": subject_correct / len(subject_expected) if subject_expected else 1.0,
        "eventFields": event_fields,
        "eventCompleteness": complete_count / len(event_matches) if event_matches else 1.0,
        "summaryQualifiedRate": summary_qualified / len(event_matches) if event_matches else 1.0,
        "irrelevantInfoRate": noise_count / len(event_matches) if event_matches else 0.0,
        "duplicateRate": duplicate_models / event_model_total if event_model_total else 0.0,
        "splitRate": len(split_gold) / len(event_gold_indices) if event_gold_indices else 0.0,
        "mergeAccuracy": merge_correct / len(merge_groups) if merge_groups else 1.0,
        "mergeGroupTotal": len(merge_groups),
        "verbalDescriptionAccuracy": verbal_description_accuracy,
        "documentaryRelationAccuracy": documentary_relation_accuracy / documentary_relation_total if documentary_relation_total else 1.0,
        "inputLimited": limited,
        "reviewRequired": sum(1 for row in matches if row["reviewRequired"]),
    }
    return {"summary": summary, "details": details, "model": parsed, "threshold": loose_threshold}


def apply_review_decisions(result, decisions):
    """Apply human accept/reject decisions and recalculate headline metrics."""
    details = result.get("details") or []
    for key, decision in (decisions or {}).items():
        try:
            row = details[int(key)]
        except (ValueError, TypeError, IndexError):
            continue
        if row.get("status") != "matched" or decision not in {"accept", "reject", ""}:
            continue
        if decision:
            row["reviewDecision"] = decision
            row["reviewRequired"] = False
        else:
            row.pop("reviewDecision", None)

    key_map = {"event": "event", "verbal_issue": "verbalIssue", "documentary_issue": "documentaryIssue"}
    for kind, summary_key in key_map.items():
        metric = result["summary"][summary_key]
        accepted = [row for row in details if row.get("status") == "matched" and (row.get("gold") or {}).get("kind") == kind and row.get("reviewDecision") != "reject"]
        tp = len(accepted)
        precision = tp / metric["modelTotal"] if metric["modelTotal"] else (1.0 if not metric["goldTotal"] else 0.0)
        recall = tp / metric["goldTotal"] if metric["goldTotal"] else 1.0
        metric.update({
            "matched": tp,
            "strictMatched": sum(1 for row in accepted if row.get("strict")),
            "precision": precision,
            "recall": recall,
            "f1": 2 * precision * recall / (precision + recall) if precision + recall else 0.0,
        })
    accepted_events = [row for row in details if row.get("status") == "matched" and (row.get("gold") or {}).get("kind") == "event" and row.get("reviewDecision") != "reject" and _norm((row.get("gold") or {}).get("subject"))]
    result["summary"]["subjectAccuracy"] = (
        sum(_similarity(row["gold"].get("subject"), row["model"].get("subject")) >= 0.90 for row in accepted_events) / len(accepted_events)
        if accepted_events else 1.0
    )
    all_accepted_events = [row for row in details if row.get("status") == "matched" and (row.get("gold") or {}).get("kind") == "event" and row.get("reviewDecision") != "reject"]
    for field in ("subject", "behavior", "eventTime", "amount", "location", "relatedObject"):
        applicable = [row for row in all_accepted_events if _norm(row["gold"].get(field))]
        correct = sum(_similarity(row["gold"].get(field), row["model"].get(field)) >= .90 for row in applicable)
        result["summary"].setdefault("eventFields", {})[field] = {"correct": correct, "total": len(applicable), "accuracy": correct / len(applicable) if applicable else 1.0}
    complete = summary_ok = noise = 0
    for row in all_accepted_events:
        gold, model = row["gold"], row["model"]
        required = [field for field in ("subject", "behavior", "eventTime", "amount", "location", "relatedObject") if _norm(gold.get(field))]
        complete += all(_norm(model.get(field)) for field in required)
        gold_summary, model_summary = _text(gold.get("summary")), _text(model.get("summary"))
        summary_ok += _similarity(gold_summary, model_summary) >= .60 and len(model_summary) <= max(40, len(gold_summary) * 3)
        noise_tokens = [part.strip() for part in re.split(r"[|,，;；]", _text(gold.get("irrelevantInfo"))) if part.strip()]
        noise += any(token in model_summary for token in noise_tokens) or bool(re.search(r"第?\s*\d+\s*页|被谈话人[:：]", model_summary))
    result["summary"]["eventCompleteness"] = complete / len(all_accepted_events) if all_accepted_events else 1.0
    result["summary"]["summaryQualifiedRate"] = summary_ok / len(all_accepted_events) if all_accepted_events else 1.0
    result["summary"]["irrelevantInfoRate"] = noise / len(all_accepted_events) if all_accepted_events else 0.0
    accepted_verbal = [row for row in details if row.get("status") == "matched" and (row.get("gold") or {}).get("kind") == "verbal_issue" and row.get("reviewDecision") != "reject"]
    result["summary"]["verbalDescriptionAccuracy"] = sum(_similarity(row["gold"].get("summary"), row["model"].get("summary")) >= .82 for row in accepted_verbal) / len(accepted_verbal) if accepted_verbal else 1.0
    accepted_documentary = [row for row in details if row.get("status") == "matched" and (row.get("gold") or {}).get("kind") == "documentary_issue" and row.get("reviewDecision") != "reject" and row["gold"].get("relatedFiles")]
    result["summary"]["documentaryRelationAccuracy"] = sum(_similarity("|".join(row["gold"].get("relatedFiles") or []), "|".join(row["model"].get("relatedFiles") or [])) >= .75 for row in accepted_documentary) / len(accepted_documentary) if accepted_documentary else 1.0
    result["summary"]["reviewRequired"] = sum(
        1 for row in details if row.get("status") == "matched" and row.get("reviewRequired") and not row.get("reviewDecision")
    )
    return result


class EvidenceEvaluationService:
    def __init__(self, db_path, output_dir):
        self.db_path = Path(db_path)
        self.output_dir = Path(output_dir)

    def _connect(self):
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        return conn

    def init_db(self):
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        with self._connect() as conn:
            conn.executescript("""
                CREATE TABLE IF NOT EXISTS evidence_datasets (
                    id TEXT PRIMARY KEY, name TEXT NOT NULL, version TEXT NOT NULL,
                    description TEXT NOT NULL DEFAULT '', status TEXT NOT NULL DEFAULT 'draft',
                    created_at TEXT NOT NULL, updated_at TEXT NOT NULL
                );
                CREATE TABLE IF NOT EXISTS evidence_gold (
                    id TEXT PRIMARY KEY, dataset_id TEXT NOT NULL, position INTEGER NOT NULL,
                    content TEXT NOT NULL, created_at TEXT NOT NULL, updated_at TEXT NOT NULL
                );
                CREATE INDEX IF NOT EXISTS idx_evidence_gold_dataset ON evidence_gold(dataset_id, position);
                CREATE TABLE IF NOT EXISTS evidence_runs (
                    id TEXT PRIMARY KEY, dataset_id TEXT NOT NULL, name TEXT NOT NULL,
                    model_version TEXT NOT NULL DEFAULT '', prompt_version TEXT NOT NULL DEFAULT '',
                    system_version TEXT NOT NULL DEFAULT '', source TEXT NOT NULL,
                    result TEXT NOT NULL, created_at TEXT NOT NULL
                );
                CREATE INDEX IF NOT EXISTS idx_evidence_runs_dataset ON evidence_runs(dataset_id, created_at);
            """)

    def list_datasets(self):
        with self._connect() as conn:
            rows = conn.execute("""
                SELECT d.*, COUNT(DISTINCT g.id) gold_count, COUNT(DISTINCT r.id) run_count
                FROM evidence_datasets d
                LEFT JOIN evidence_gold g ON g.dataset_id=d.id
                LEFT JOIN evidence_runs r ON r.dataset_id=d.id
                GROUP BY d.id ORDER BY d.updated_at DESC
            """).fetchall()
        return [{"id": row["id"], "name": row["name"], "version": row["version"],
                 "description": row["description"], "status": row["status"],
                 "goldCount": row["gold_count"], "runCount": row["run_count"],
                 "createdAt": row["created_at"], "updatedAt": row["updated_at"]} for row in rows]

    def get_dataset(self, dataset_id):
        with self._connect() as conn:
            row = conn.execute("SELECT * FROM evidence_datasets WHERE id=?", (dataset_id,)).fetchone()
            if not row:
                raise ValueError("评测集不存在。")
            gold = [json.loads(item["content"]) for item in conn.execute(
                "SELECT content FROM evidence_gold WHERE dataset_id=? ORDER BY position", (dataset_id,)
            )]
            runs = conn.execute("SELECT id,name,model_version,prompt_version,system_version,created_at,result FROM evidence_runs WHERE dataset_id=? ORDER BY created_at DESC", (dataset_id,)).fetchall()
        return {
            "id": row["id"], "name": row["name"], "version": row["version"],
            "description": row["description"], "status": row["status"],
            "createdAt": row["created_at"], "updatedAt": row["updated_at"], "gold": gold,
            "runs": [{"id": run["id"], "name": run["name"], "modelVersion": run["model_version"],
                      "promptVersion": run["prompt_version"], "systemVersion": run["system_version"],
                      "createdAt": run["created_at"], "summary": json.loads(run["result"])["summary"]} for run in runs],
        }

    def save_dataset(self, payload):
        dataset_id = _text(payload.get("id")) or uuid.uuid4().hex
        name = _text(payload.get("name"))
        if not name:
            raise ValueError("请填写评测集名称。")
        now = time.strftime("%Y-%m-%d %H:%M:%S")
        with self._connect() as conn:
            existing = conn.execute("SELECT status,created_at FROM evidence_datasets WHERE id=?", (dataset_id,)).fetchone()
            if existing and existing["status"] == "frozen" and payload.get("status") != "frozen":
                raise ValueError("已冻结评测集不能解除冻结，请复制为新版本。")
            conn.execute("""
                INSERT INTO evidence_datasets(id,name,version,description,status,created_at,updated_at)
                VALUES(?,?,?,?,?,?,?)
                ON CONFLICT(id) DO UPDATE SET name=excluded.name,version=excluded.version,
                    description=excluded.description,status=excluded.status,updated_at=excluded.updated_at
            """, (dataset_id, name, _text(payload.get("version")) or "v1.0",
                  _text(payload.get("description")), _text(payload.get("status")) or "draft",
                  existing["created_at"] if existing else now, now))
        return self.get_dataset(dataset_id)

    def delete_dataset(self, dataset_id):
        dataset_id = _text(dataset_id)
        if not dataset_id:
            raise ValueError("缺少评测集编号。")
        with self._connect() as conn:
            dataset = conn.execute("SELECT name FROM evidence_datasets WHERE id=?", (dataset_id,)).fetchone()
            if not dataset:
                raise ValueError("评测集不存在。")
            gold_count = conn.execute("SELECT COUNT(*) FROM evidence_gold WHERE dataset_id=?", (dataset_id,)).fetchone()[0]
            run_count = conn.execute("SELECT COUNT(*) FROM evidence_runs WHERE dataset_id=?", (dataset_id,)).fetchone()[0]
            conn.execute("DELETE FROM evidence_runs WHERE dataset_id=?", (dataset_id,))
            conn.execute("DELETE FROM evidence_gold WHERE dataset_id=?", (dataset_id,))
            conn.execute("DELETE FROM evidence_datasets WHERE id=?", (dataset_id,))
        return {"deleted": True, "id": dataset_id, "name": dataset["name"],
                "goldCount": gold_count, "runCount": run_count}

    def save_gold(self, dataset_id, rows):
        dataset = self.get_dataset(dataset_id)
        if dataset["status"] == "frozen":
            raise ValueError("评测集已冻结，不能修改黄金标注。")
        normalized = []
        valid_kinds = {"event", "verbal_issue", "documentary_issue"}
        for index, row in enumerate(rows or []):
            kind = KIND_ALIASES.get(_text(row.get("kind")), _text(row.get("kind")))
            if kind not in valid_kinds:
                raise ValueError(f"第{index + 1}条标注类型无效。")
            item = {key: _text(row.get(key)) for key in (
                "id", "caseId", "fileId", "itemId", "itemName", "pageNo", "sourceText",
                "subject", "behavior", "eventTime", "amount", "location", "relatedObject",
                "summary", "evidenceText", "eventClusterId", "mergeRelation",
                "expectedConclusion", "split", "annotationStatus", "notes",
                "issueScene", "documentStatus", "limitationReason", "irrelevantInfo",
            )}
            item.update({"id": item["id"] or uuid.uuid4().hex, "kind": kind,
                         "relatedFiles": row.get("relatedFiles") if isinstance(row.get("relatedFiles"), list) else [part.strip() for part in re.split(r"[|；;]", _text(row.get("relatedFiles"))) if part.strip()],
                         "inputLimited": row.get("inputLimited") if isinstance(row.get("inputLimited"), bool) else _bool(row.get("inputLimited"))})
            if not item["summary"]:
                raise ValueError(f"第{index + 1}条标注缺少标准事件或问题。")
            normalized.append(item)
        now = time.strftime("%Y-%m-%d %H:%M:%S")
        with self._connect() as conn:
            conn.execute("DELETE FROM evidence_gold WHERE dataset_id=?", (dataset_id,))
            conn.executemany("INSERT INTO evidence_gold(id,dataset_id,position,content,created_at,updated_at) VALUES(?,?,?,?,?,?)",
                             [(item["id"], dataset_id, pos, json.dumps(item, ensure_ascii=False), now, now) for pos, item in enumerate(normalized)])
            conn.execute("UPDATE evidence_datasets SET updated_at=? WHERE id=?", (now, dataset_id))
        return self.get_dataset(dataset_id)

    def import_gold_excel(self, dataset_id, encoded):
        raw = base64.b64decode(encoded)
        wb = openpyxl.load_workbook(BytesIO(raw), data_only=True)
        ws = wb["黄金标注"] if "黄金标注" in wb.sheetnames else wb.active
        headers = {_text(cell.value): index for index, cell in enumerate(ws[1])}
        optional_titles = {"问题场景", "书证存在状态", "匹配限制原因", "无关信息"}
        missing = [title for _, title, _ in GOLD_COLUMNS if title not in headers and title not in optional_titles]
        if missing:
            raise ValueError("Excel缺少列：" + "、".join(missing))
        rows = []
        for values in ws.iter_rows(min_row=2, values_only=True):
            if not any(value not in (None, "") for value in values):
                continue
            raw_row = {key: values[headers[title]] if title in headers else "" for key, title, _ in GOLD_COLUMNS}
            rows.append({
                "kind": KIND_ALIASES.get(_text(raw_row["kind"]), _text(raw_row["kind"])), "caseId": _text(raw_row["case_id"]),
                "fileId": _text(raw_row["file_id"]), "itemId": _text(raw_row["item_id"]),
                "itemName": _text(raw_row["item_name"]), "pageNo": _text(raw_row["page_no"]),
                "sourceText": _text(raw_row["source_text"]), "subject": _text(raw_row["subject"]),
                "behavior": _text(raw_row["behavior"]), "eventTime": _text(raw_row["event_time"]),
                "amount": _text(raw_row["amount"]), "location": _text(raw_row["location"]),
                "relatedObject": _text(raw_row["related_object"]), "summary": _text(raw_row["summary"]),
                "evidenceText": _text(raw_row["evidence_text"]), "eventClusterId": _text(raw_row["event_cluster_id"]),
                "mergeRelation": _text(raw_row["merge_relation"]), "expectedConclusion": _text(raw_row["expected_conclusion"]),
                "issueScene": _text(raw_row["issue_scene"]), "documentStatus": _text(raw_row["document_status"]),
                "relatedFiles": _text(raw_row["related_files"]), "inputLimited": _bool(raw_row["input_limited"]),
                "limitationReason": _text(raw_row["limitation_reason"]), "irrelevantInfo": _text(raw_row["irrelevant_info"]),
                "split": _text(raw_row["split"]), "annotationStatus": _text(raw_row["annotation_status"]),
                "notes": _text(raw_row["notes"]),
            })
        return self.save_gold(dataset_id, rows)

    def run(self, payload):
        dataset = self.get_dataset(_text(payload.get("datasetId")))
        if not dataset["gold"]:
            raise ValueError("当前评测集还没有黄金标注。")
        source = payload.get("source")
        if not isinstance(source, dict):
            raise ValueError("系统结果必须是JSON对象。")
        result = evaluate(dataset["gold"], source, float(payload.get("threshold") or 0.58))
        run_id, now = uuid.uuid4().hex, time.strftime("%Y-%m-%d %H:%M:%S")
        with self._connect() as conn:
            conn.execute("INSERT INTO evidence_runs(id,dataset_id,name,model_version,prompt_version,system_version,source,result,created_at) VALUES(?,?,?,?,?,?,?,?,?)",
                         (run_id, dataset["id"], _text(payload.get("name")) or f"评测运行 {now}",
                          _text(payload.get("modelVersion")), _text(payload.get("promptVersion")),
                          _text(payload.get("systemVersion")), json.dumps(source, ensure_ascii=False),
                          json.dumps(result, ensure_ascii=False), now))
        return {"id": run_id, "name": _text(payload.get("name")) or f"评测运行 {now}", "createdAt": now, **result}

    def get_run(self, run_id):
        with self._connect() as conn:
            row = conn.execute("SELECT * FROM evidence_runs WHERE id=?", (run_id,)).fetchone()
        if not row:
            raise ValueError("评测记录不存在。")
        return {"id": row["id"], "datasetId": row["dataset_id"], "name": row["name"],
                "modelVersion": row["model_version"], "promptVersion": row["prompt_version"],
                "systemVersion": row["system_version"], "createdAt": row["created_at"], **json.loads(row["result"])}

    def review_run(self, run_id, decisions):
        with self._connect() as conn:
            row = conn.execute("SELECT result FROM evidence_runs WHERE id=?", (run_id,)).fetchone()
            if not row:
                raise ValueError("评测记录不存在。")
            result = apply_review_decisions(json.loads(row["result"]), decisions)
            conn.execute("UPDATE evidence_runs SET result=? WHERE id=?", (json.dumps(result, ensure_ascii=False), run_id))
        return self.get_run(run_id)

    def template(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "黄金标注"
        ws.append([title for _, title, _ in GOLD_COLUMNS])
        ws.append(["事件", "示例案件", "file-1", "item-1", "谈话笔录", "12", "原始问答文本", "被调查人", "兼职取酬", "2024年", "10000元", "", "某公司", "被调查人在某公司兼职并取得报酬", "对应原文", "E001", "独立", "", "兼职取酬", "", "", "否", "", "页码、被谈话人", "开发集", "已确认", "示例行，导入前可删除"])
        note = wb.create_sheet("填写说明")
        note.append(["字段", "说明"])
        for _, title, description in GOLD_COLUMNS:
            note.append([title, description])
        for sheet in wb.worksheets:
            sheet.freeze_panes = "A2"
            for cell in sheet[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="2563EB")
            for column in range(1, sheet.max_column + 1):
                sheet.column_dimensions[get_column_letter(column)].width = 18
        stream = BytesIO()
        wb.save(stream)
        return base64.b64encode(stream.getvalue()).decode("ascii")

    def export_run(self, run_id):
        run = self.get_run(run_id)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        path = self.output_dir / f"言证书证评测报告_{time.strftime('%Y%m%d_%H%M%S')}.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "指标汇总"
        ws.append(["评测名称", run["name"]])
        ws.append(["运行时间", run["createdAt"]])
        ws.append(["模型版本", run["modelVersion"]])
        ws.append(["提示词版本", run["promptVersion"]])
        ws.append(["系统版本", run["systemVersion"]])
        ws.append([])
        ws.append(["评测对象", "黄金数", "模型数", "宽松匹配", "严格匹配", "精确率", "召回率", "F1"])
        for key, label in (("event", "事件抽取"), ("verbalIssue", "言证问题"), ("documentaryIssue", "书证问题")):
            item = run["summary"][key]
            ws.append([label, item["goldTotal"], item["modelTotal"], item["matched"], item["strictMatched"], item["precision"], item["recall"], item["f1"]])
        ws.append([])
        ws.append(["主体归属正确率", run["summary"]["subjectAccuracy"]])
        ws.append(["事件完整率", run["summary"].get("eventCompleteness", 0)])
        ws.append(["摘要合格率", run["summary"].get("summaryQualifiedRate", 0)])
        ws.append(["无关信息率", run["summary"].get("irrelevantInfoRate", 0)])
        ws.append(["事件重复率", run["summary"]["duplicateRate"]])
        ws.append(["事件拆分率", run["summary"]["splitRate"]])
        ws.append(["事件合并准确率", run["summary"].get("mergeAccuracy", 0)])
        ws.append(["言证描述正确率", run["summary"].get("verbalDescriptionAccuracy", 0)])
        ws.append(["书证关联准确率", run["summary"].get("documentaryRelationAccuracy", 0)])
        ws.append(["输入范围受限", run["summary"]["inputLimited"]])
        ws.append(["待人工复核", run["summary"]["reviewRequired"]])
        ws.append([])
        ws.append(["事件字段", "正确数", "匹配数", "准确率"])
        field_labels = {"subject": "主体", "behavior": "行为", "eventTime": "时间", "amount": "金额",
                        "location": "地点", "relatedObject": "关联对象"}
        for key, metric in run["summary"].get("eventFields", {}).items():
            ws.append([field_labels.get(key, key), metric.get("correct", 0), metric.get("total", 0), metric.get("accuracy", 0)])
        detail = wb.create_sheet("匹配明细")
        detail.append(["状态", "类型", "案件", "材料", "页码", "黄金答案", "系统答案", "匹配分", "严格匹配", "备注"])
        for row in run["details"]:
            gold, model = row.get("gold") or {}, row.get("model") or {}
            detail.append([row["status"], gold.get("kind") or model.get("kind"), gold.get("caseId", ""),
                           gold.get("itemName") or model.get("itemName", ""), gold.get("pageNo", ""),
                           gold.get("summary", ""), model.get("summary", ""), row.get("score", ""),
                           "是" if row.get("strict") else "否", gold.get("notes", "")])
        for sheet in wb.worksheets:
            for row in sheet.iter_rows():
                for cell in row:
                    cell.alignment = Alignment(vertical="top", wrap_text=True)
            for column in range(1, sheet.max_column + 1):
                sheet.column_dimensions[get_column_letter(column)].width = 20
        wb.save(path)
        return path
