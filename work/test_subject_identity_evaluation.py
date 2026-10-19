import base64
import sys
import tempfile
import unittest
from io import BytesIO
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent))

from subject_identity_evaluation import SubjectIdentityEvaluationService, evaluate, parse_output


GOLD = {
    "cases": [{"case_id": "C001", "first_violation_date": "2022-01-01"}],
    "files": [{"case_id": "C001", "file_id": "F001", "item_id": "I001", "file_name": "任职文件1", "expected_file_type": "任职文件", "page": "3", "should_process": "是"}],
    "records": [{"record_id": "R001", "case_id": "C001", "file_id": "F001", "item_id": "I001", "file_name": "任职文件1", "job": "XX局局长", "start_date": "2020-01", "end_date": "2022-06", "job_type": "职务", "file_type": "任职文件", "page": "3"}],
    "groups": [{"group_id": "G001", "record_id": "R001", "case_id": "C001", "normalized_job": "XX局局长", "expected_start": "2020-01", "expected_end": "2022-06", "relation": "任职", "should_merge": "是"}],
    "identities": [{"identity_id": "ID001", "case_id": "C001", "identity_type": "特殊身份", "identity": "人大代表", "case_input": "人大代表", "report_input": "", "material_file_id": "F002", "material_item_id": "I002", "material_name": "代表证明", "page": "1"}],
    "rules": [{"rule_id": "BR001", "case_id": "C001", "rule_type": "过滤", "expected_result": "过滤或标记"}],
}

MODEL = {
    "caseId": "C001",
    "fileClassifications": [{"fileId": "F001", "itemId": "I001", "fileName": "任职文件1", "fileType": "任职文件", "page": "3"}],
    "positions": [{"id": "M1", "fileId": "F001", "itemId": "I001", "job": "XX局局长", "startDate": "2020-01", "endDate": "2022-06", "jobType": "职务", "fileType": "任职文件", "page": "3"}],
    "positionGroups": [{"groupId": "MG1", "job": "XX局局长", "startDate": "2020-01", "endDate": "2022-06", "records": [{"fileId": "F001", "itemId": "I001", "job": "XX局局长", "startDate": "2020-01", "endDate": "2022-06"}]}],
    "specialIdentities": [{"identity": "人大代表", "caseInput": "人大代表", "materials": [{"fileId": "F002", "itemId": "I002", "material": "代表证明", "page": "1"}]}],
    "businessRules": [{"ruleId": "BR001", "result": "过滤或标记"}],
}


class SubjectEvaluationTest(unittest.TestCase):
    def test_parses_four_stages(self):
        parsed = parse_output({"data": MODEL})
        self.assertEqual(1, len(parsed["files"]))
        self.assertGreaterEqual(len(parsed["records"]), 1)
        self.assertEqual(1, len(parsed["groups"]))
        self.assertEqual(1, len(parsed["identities"]))

    def test_evaluates_four_stages_and_fields(self):
        result = evaluate(GOLD, MODEL)
        self.assertEqual(1, result["summary"]["fileClassification"]["matched"])
        self.assertEqual(1, result["summary"]["recordExtraction"]["matched"])
        self.assertEqual(1, result["summary"]["alignment"]["matched"])
        self.assertEqual(1, result["summary"]["identity"]["matched"])
        self.assertEqual(1.0, result["summary"]["materialAccuracy"])
        self.assertEqual(1.0, result["summary"]["businessRules"]["passRate"])
        self.assertEqual(1.0, result["summary"]["fields"]["job"]["accuracy"])

    def test_parses_actual_identity_endpoint_shape(self):
        payload = {"code": 0, "data": {"caseId": "C001", "positions": [{
            "formStartTime": "2020-01", "formEndTime": "2022-06", "formPosition": "XX局局长",
            "fileId": "F001", "pageNum": 3, "itemName": "干部履历表",
            "appointmentDocs": [{"time": "2020-01", "position": "XX局局长", "itemName": "任职文件", "fileId": "F003", "pageNum": 2, "actionType": "任职"}],
            "divisionDocs": [{"time": "2021-01", "position": "分管财务", "itemName": "分工文件", "fileId": "F004", "pageNum": 1}],
        }], "ranks": [{"formStartTime": "2021-01", "formRank": "一级主任科员", "promotionDocs": []}],
        "specialIdentities": [{"creationSelection": "人大代表", "investigateReport": "人大代表", "proofDocument": [{"fileId": "F002", "pageNum": 1, "itemName": "代表证明"}]}],
        "otherIdentities": []}}
        parsed = parse_output(payload)
        self.assertEqual(4, len(parsed["records"]))
        self.assertEqual(2, len(parsed["groups"]))
        self.assertEqual("人大代表", parsed["identities"][0]["identity"])
        self.assertEqual("F002", parsed["identities"][0]["materials"][0]["fileId"])


class SubjectServiceTest(unittest.TestCase):
    def test_dataset_template_import_run_export_delete(self):
        with tempfile.TemporaryDirectory() as folder:
            service = SubjectIdentityEvaluationService(Path(folder) / "db.sqlite", Path(folder) / "out")
            service.init_db()
            dataset = service.save_dataset({"name": "主体身份测试", "version": "v1"})
            encoded = service.template()
            self.assertGreater(len(base64.b64decode(encoded)), 1000)
            workbook = openpyxl.load_workbook(BytesIO(base64.b64decode(encoded)), read_only=True)
            self.assertEqual("案件编号", workbook["案件信息"]["A1"].value)
            self.assertEqual("标准职位或身份", workbook["职位抽取"]["F1"].value)
            self.assertEqual("是否应进入阅卷问题", workbook["业务规则"]["G1"].value)
            dataset = service.import_gold(dataset["id"], encoded)
            self.assertEqual(1, len(dataset["gold"]["records"]))
            run = service.run({"datasetId": dataset["id"], "source": MODEL, "name": "回归"})
            self.assertTrue(service.export_run(run["id"]).exists())
            deleted = service.delete_dataset(dataset["id"])
            self.assertEqual(1, deleted["runCount"])
            self.assertEqual([], service.list_datasets())


if __name__ == "__main__":
    unittest.main()
