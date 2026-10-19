import base64
import tempfile
import unittest
from io import BytesIO
from pathlib import Path
import sys

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent))

from evidence_evaluation import EvidenceEvaluationService, evaluate, flatten_model_result


def gold_event(**changes):
    row = {
        "id": "g1", "kind": "event", "caseId": "case-1", "fileId": "f1",
        "itemId": "i1", "itemName": "谈话笔录", "pageNo": "12",
        "sourceText": "张某在某公司兼职并取得报酬一万元。", "subject": "张某",
        "behavior": "兼职取酬", "eventTime": "2024年", "amount": "10000元",
        "location": "", "relatedObject": "某公司",
        "summary": "张某在某公司兼职并取得报酬一万元",
        "evidenceText": "张某在某公司兼职并取得报酬一万元。",
        "eventClusterId": "E1", "mergeRelation": "独立", "expectedConclusion": "",
        "issueScene": "兼职取酬", "documentStatus": "", "relatedFiles": [],
        "inputLimited": False, "limitationReason": "", "irrelevantInfo": "页码",
        "split": "开发集", "annotationStatus": "已确认", "notes": "",
    }
    row.update(changes)
    return row


class EvaluationTest(unittest.TestCase):
    def test_flattens_events_and_evidence_issues(self):
        payload = {"data": {
            "extract_transcript_events": [{
                "file_id": "f1", "item_id": "i1", "item_name": "谈话笔录",
                "events": [{"subject": "张某", "event_type": "兼职取酬", "description": "张某兼职取酬"}],
            }],
            "issues": [
                {"type": "VERBAL_EVIDENCE", "description": "两份笔录金额矛盾"},
                {"type": "DOCUMENTARY_EVIDENCE", "description": "缺少转账凭证", "related_files": ["转账凭证"]},
                {"type": "OTHER", "description": "不在评测范围"},
            ],
        }}
        result = flatten_model_result(payload)
        self.assertEqual(1, len(result["events"]))
        self.assertEqual(2, len(result["issues"]))
        self.assertEqual("f1", result["events"][0]["fileId"])

    def test_evaluates_three_channels_and_excludes_input_limit(self):
        gold = [
            gold_event(),
            {**gold_event(id="g2", kind="verbal_issue", summary="两份笔录金额矛盾", behavior="金额矛盾")},
            {**gold_event(id="g3", kind="documentary_issue", summary="缺少转账凭证", relatedFiles=["转账凭证"], inputLimited=True)},
        ]
        payload = {"extract_transcript_events": [{
            "file_id": "f1", "item_id": "i1", "item_name": "谈话笔录",
            "events": [{"subject": "张某", "event_type": "兼职取酬", "time": "2024年", "amount": "10000元", "object": "某公司", "description": "张某在某公司兼职并取得报酬一万元"}],
        }], "issues": [{"type": "VERBAL_EVIDENCE", "description": "两份笔录金额矛盾"}]}
        result = evaluate(gold, payload)
        self.assertEqual(1, result["summary"]["event"]["matched"])
        self.assertEqual(1, result["summary"]["verbalIssue"]["matched"])
        self.assertEqual(0, result["summary"]["documentaryIssue"]["goldTotal"])
        self.assertEqual(1, result["summary"]["inputLimited"])
        self.assertIn("eventFields", result["summary"])
        self.assertIn("eventCompleteness", result["summary"])
        self.assertIn("summaryQualifiedRate", result["summary"])
        self.assertIn("irrelevantInfoRate", result["summary"])
        self.assertIn("mergeAccuracy", result["summary"])


class ServiceTest(unittest.TestCase):
    def test_dataset_gold_run_and_export_lifecycle(self):
        with tempfile.TemporaryDirectory() as folder:
            service = EvidenceEvaluationService(Path(folder) / "test.db", Path(folder) / "out")
            service.init_db()
            dataset = service.save_dataset({"name": "测试集", "version": "v1"})
            template_bytes = base64.b64decode(service.template())
            workbook = openpyxl.load_workbook(BytesIO(template_bytes), data_only=True)
            headers = [cell.value for cell in workbook["黄金标注"][1]]
            self.assertEqual("事件", workbook["黄金标注"]["A2"].value)
            for header in ("问题场景", "书证存在状态", "匹配限制原因", "无关信息"):
                self.assertIn(header, headers)
            imported = service.import_gold_excel(dataset["id"], base64.b64encode(template_bytes).decode("ascii"))
            self.assertEqual("event", imported["gold"][0]["kind"])
            saved = service.save_gold(dataset["id"], [gold_event()])
            self.assertEqual(1, len(saved["gold"]))
            run = service.run({
                "datasetId": dataset["id"], "name": "测试运行",
                "source": {"extract_transcript_events": [{"events": [{"subject": "张某", "event_type": "兼职取酬", "description": "张某在某公司兼职并取得报酬一万元"}]}]},
            })
            self.assertEqual(1, run["summary"]["event"]["matched"])
            reviewed = service.review_run(run["id"], {"0": "reject"})
            self.assertEqual(0, reviewed["summary"]["event"]["matched"])
            self.assertEqual(0.0, reviewed["summary"]["event"]["f1"])
            self.assertTrue(service.export_run(run["id"]).exists())
            frozen = service.save_dataset({"id": dataset["id"], "name": "测试集", "version": "v1", "status": "frozen"})
            self.assertEqual("frozen", frozen["status"])
            with self.assertRaises(ValueError):
                service.save_gold(dataset["id"], [gold_event(summary="修改")])
            deleted = service.delete_dataset(dataset["id"])
            self.assertTrue(deleted["deleted"])
            self.assertEqual(1, deleted["goldCount"])
            self.assertEqual(1, deleted["runCount"])
            self.assertEqual([], service.list_datasets())


if __name__ == "__main__":
    unittest.main()
