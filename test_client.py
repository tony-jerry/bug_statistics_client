"""不连接线上服务的核心逻辑测试。"""

from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch

from openpyxl import Workbook

from analytics import filter_bugs, introducer_name, introducer_summary, summary_metrics
from credential_store import load_credentials, save_credentials
from crm_client import (
    CRM_SEARCH_PAGE_SIZE,
    BugDetail,
    BugDetailField,
    build_crm_bug_url,
    build_crm_search_payload,
    extract_detail_fields,
    extract_crm_object_identity,
    encrypt_crm_password,
    normalize_crm_bug,
)
from detail_utils import (
    extract_image_sources,
    rich_content_segments,
    rich_text_to_plain,
)
from retrospective import (
    build_retrospective_markdown,
    retrospective_filename,
)
from workload_client import (
    ALL_GROUP_MEMBERS,
    SERVER_WORKLOAD_DEFINITION,
    UI_WORKLOAD_DEFINITION,
    Developer,
    WorkloadApiClient,
    WorkloadContext,
    allocate_ui_workload,
    allocate_server_workload,
    build_workload_preview,
    calculate_workload_hours,
    read_workload_excel,
)


SAMPLE = [
    {
        "id": "Bug001",
        "createdByName": "危国",
        "createdBy": "weig",
        "severity": "2",
        "status": "关闭",
        "title": "导出失败",
        "module": "公共模块",
    },
    {
        "id": "Bug002",
        "createdByName": "危国",
        "severity": "3",
        "status": "处理中",
        "title": "样式异常",
        "module": "表单管理",
    },
    {
        "id": "Bug003",
        "createdByName": "张三",
        "severity": "2",
        "status": "关闭",
        "title": "查询失败",
        "module": "公共模块",
    },
]


def test_analytics() -> None:
    assert introducer_name(SAMPLE[0]) == "危国"
    filtered = filter_bugs(SAMPLE, introducer="危国")
    assert [item["id"] for item in filtered] == ["Bug001", "Bug002"]
    assert len(filter_bugs(SAMPLE, keyword="样式")) == 1

    metrics = summary_metrics(filtered)
    assert metrics["total"] == 2
    assert metrics["severity_2"] == 1
    assert metrics["severity_3"] == 1
    assert metrics["closed"] == 1
    assert metrics["open"] == 1

    rows = introducer_summary(SAMPLE)
    assert rows[0]["introducer"] == "危国"
    assert rows[0]["total"] == 2


def test_crm_helpers() -> None:
    assert (
        encrypt_crm_password("example-password")
        == "UU9iMVo0MjFiOEFpREtXWDY0YUE4YU8yOXRiWk9Rb1N2azNCTmVNS3I3QT0="
    )
    url = (
        'https://crm.example/?objOrginalParam=%7B%22oid%22%3A%22123%22%2C'
        '%22otype%22%3A%22ty.inteplm.ipd.CTyDefect%22%7D'
    )
    assert extract_crm_object_identity(url) == (
        "123",
        "ty.inteplm.ipd.CTyDefect",
    )
    assert build_crm_bug_url(
        "https://crm.example/",
        "123",
        "ty.inteplm.ipd.CTyDefect",
        "Bug 001",
    ) == (
        "https://crm.example/#/homePage/defectObjForm?"
        "otype=ty.inteplm.ipd.CTyDefect&oid=123&label=Bug+001"
    )
    assert rich_text_to_plain("<p>原因：时序问题</p><p><img src='x'></p>") == (
        "原因：时序问题\n[图片]"
    )
    rich = "<p>验证截图：</p><p><img src='/rest/image?a=1&amp;b=2'></p>"
    assert extract_image_sources(rich) == ["/rest/image?a=1&b=2"]
    assert rich_content_segments(rich) == [
        ("text", "验证截图：\n"),
        ("image", "/rest/image?a=1&b=2"),
    ]


def test_credential_store() -> None:
    with TemporaryDirectory() as directory:
        path = Path(directory) / "credentials.dat"
        assert save_credentials(path, "tester", "performance-pass", "crm-pass")
        encrypted = path.read_bytes()
        assert b"performance-pass" not in encrypted
        assert b"crm-pass" not in encrypted
        assert load_credentials(path) == {
            "username": "tester",
            "password": "performance-pass",
            "crm_password": "crm-pass",
        }


def test_all_detail_fields() -> None:
    form_data = {
        "name": "示例缺陷",
        "owner": {"name": "T0001", "fullname": "测试用户"},
        "ibaAttribute": {
            "phase": ["SIT"],
            "IsAutoDetected": ["0"],
            "DTS_TestResult": ["<p>测试通过</p>"],
        },
    }
    template = {
        "widgetList": [
            {
                "type": "input",
                "options": {"realProps": "name", "label": "标题"},
            },
            {
                "type": "select-user",
                "options": {"realProps": "owner", "label": "负责人"},
            },
            {
                "type": "radio",
                "options": {
                    "realProps": "phase",
                    "label": "阶段",
                    "optionItems": [{"value": "SIT", "label": "SIT"}],
                },
            },
            {
                "type": "radio",
                "options": {
                    "realProps": "IsAutoDetected",
                    "label": "是否自动化发现",
                    "optionItems": [
                        {"value": "0", "label": "否"},
                        {"value": "1", "label": "是"},
                    ],
                },
            },
            {
                "type": "date",
                "options": {"realProps": "plandate", "label": "期望完成日期"},
            },
            {
                "type": "rich-editor",
                "options": {"realProps": "DTS_TestResult", "label": "测试结果"},
            },
        ]
    }
    fields = extract_detail_fields(form_data, template)
    assert [field.label for field in fields] == [
        "标题",
        "负责人",
        "阶段",
        "是否自动化发现",
        "期望完成日期",
        "测试结果",
    ]
    assert [field.value for field in fields[:-1]] == [
        "示例缺陷",
        "测试用户",
        "SIT",
        "否",
        "（空）",
    ]
    assert fields[-1].is_rich_text
    assert fields[-1].value == "<p>测试通过</p>"


def test_crm_search_helpers() -> None:
    count_payload = build_crm_search_payload(
        "Bug20260723003",
        1,
        need_count=True,
    )
    assert count_payload["pageNum"] == 0
    assert count_payload["pageSize"] == CRM_SEARCH_PAGE_SIZE == 50
    assert count_payload["needCount"] is True
    assert "isSetColumn" not in count_payload

    list_payload = build_crm_search_payload(
        "Bug20260723003",
        2,
        need_count=False,
    )
    assert list_payload["pageNum"] == 2
    assert list_payload["pageSize"] == 50
    assert list_payload["needCount"] is False
    assert list_payload["isSetColumn"] is False

    normalized = normalize_crm_bug(
        {
            "oid": "123",
            "otype": "ty.inteplm.ipd.CTyDefect",
            "objectNumber": "Bug001",
            "name": "回退标题",
            "lifecyclestagekeyName": "分析中",
            "creatorDisplayName": "测试用户(T0001)",
            "ibaAttrMapForSearch": {
                "827308980874936320": "2.0730",
                "773662542538473472": "公共模块 | PubSVC",
            },
            "extAttrMapForSearch": {
                "title": "缺陷标题",
                "defectCategory": "代码错误",
                "severity": "严重",
                "priorityName": "普通",
                "leader": "负责人(T0002)",
            },
        }
    )
    assert normalized["objectNumber"] == "Bug001"
    assert normalized["title"] == "缺陷标题"
    assert normalized["827308980874936320"] == "2.0730"
    assert normalized["773662542538473472"] == "公共模块 | PubSVC"
    assert normalized["_oid"] == "123"


def test_retrospective_markdown() -> None:
    fields = (
        BugDetailField(
            "description",
            "描述",
            "<p>问题描述</p>",
            True,
        ),
        BugDetailField(
            "DTS_RootCauseAnalysis",
            "根因分析",
            "<p>时序问题</p>",
            True,
        ),
        BugDetailField(
            "DTS_FixContent",
            "修复内容",
            "<p>增加状态同步</p>",
            True,
        ),
    )
    detail = BugDetail(
        workflow_key="1",
        object_oid="2",
        bug_number="Bug001",
        title="示例缺陷",
        description_html="<p>问题描述</p>",
        root_cause_html="<p>时序问题</p>",
        fix_content_html="<p>增加状态同步</p>",
        fields=fields,
    )
    markdown = build_retrospective_markdown(
        "2026-0730",
        "危国",
        [
            {
                "id": "Bug001",
                "severity": "2",
                "status": "关闭",
                "module": "公共模块 | PubSVC",
                "resolvedByName": "危国",
                "title": "示例缺陷",
            }
        ],
        {"Bug001": detail},
    )
    assert retrospective_filename("2026-0730", "危国") == "0730版本bug回溯-危国.md"
    assert "# 0730版本 bug回溯" in markdown
    assert "**周期**：0730版本SIT测试" in markdown
    assert "## 一、缺陷概况" in markdown
    assert "## 二、主要问题根因" in markdown
    assert "## 三、典型Bug深度分析" in markdown
    assert "**问题**：示例缺陷" in markdown
    assert "**根因**：时序问题" in markdown
    assert "**已修复**：增加状态同步" in markdown
    assert "## 四、后续改进" in markdown
    assert "时序问题" in markdown
    assert "增加状态同步" in markdown
    assert "**模块**：公共模块 | PubSVC" in markdown
    assert "Bug逐项回溯" not in markdown


def test_workload_import() -> None:
    allocation = allocate_ui_workload(16)
    assert allocation["reqAnalyzeWorkLoad"] == 6
    assert allocation["keyControlsCount"] == 1
    assert calculate_workload_hours(allocation) == 16

    server_allocation = allocate_server_workload(12)
    assert server_allocation["methodType"] == "全新开发"
    assert server_allocation["reqAnalyzeWorkLoad"] == 2
    assert server_allocation["codeComplexity"] == 4
    assert server_allocation["unitTestNumber"] == 6
    assert calculate_workload_hours(server_allocation) == 12
    for hours in (1, 2, 4, 8, 12, 16, 24, 32):
        generated = allocate_server_workload(hours, "普通后端任务")
        assert calculate_workload_hours(generated) == hours
        assert generated["reqAnalyzeWorkLoad"] in {0, 2, 4, 6, 8}
        assert generated["codeComplexity"] in {0, 1, 2, 4, 8, 16}

    encapsulation = allocate_server_workload(8, "接口封装")
    assert encapsulation["methodType"] == "业务封装"
    assert encapsulation["unitTestNumber"] == 4
    assert calculate_workload_hours(encapsulation) == 8
    database_task = allocate_server_workload(12, "数据库查询")
    assert database_task["isDBOperation"] is True
    assert calculate_workload_hours(database_task) == 12
    ui_debug_task = allocate_server_workload(12, "前端联调")
    assert ui_debug_task["isNeedUiDebug"] is True
    assert calculate_workload_hours(ui_debug_task) == 12
    existing_server_record = {
        "methodType": "全新开发",
        "isDBOperation": True,
        "dbAdapteNumer": 8,
        "codeComplexity": 4,
        "unitTestNumber": 8,
        "isNeedUiDebug": True,
    }
    assert calculate_workload_hours(existing_server_record) == 20.4

    workload_api = WorkloadApiClient(object())
    workload_api._post_json = lambda *args, **kwargs: {
        "data": {
            "children": [
                {
                    "workItemGroup": "开发一组",
                    "methods": [{"getUiDef": SERVER_WORKLOAD_DEFINITION}],
                },
                {
                    "workItemGroup": "开发一组-前端",
                    "methods": [{"getUiDef": UI_WORKLOAD_DEFINITION}],
                },
            ]
        }
    }
    groups = workload_api.get_groups()
    assert [group.name for group in groups] == ["开发一组", "开发一组-前端"]
    assert workload_api._definition_path("开发一组") == SERVER_WORKLOAD_DEFINITION

    with TemporaryDirectory() as directory:
        path = Path(directory) / "workload.xlsx"
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(
            [
                "需求号",
                "工作描述",
                "计划开始时间",
                "计划结束时间",
                "工时/天",
                "责任人",
            ]
        )
        sheet.append(
            [
                "SR001",
                "批量录入测试",
                "2026-08-10",
                "2026-08-11",
                2,
                "危国",
            ]
        )
        sheet.append(
            [
                "SR404",
                "不存在的需求",
                "2026-08-12",
                "2026-08-13",
                1,
                "危国",
            ]
        )
        workbook.save(path)
        workbook.close()

        items = read_workload_excel(path, "2026-0830")
        assert len(items) == 2
        assert items[0].plan_start_date == "2026-08-10"

        valid_content = path.read_bytes()
        with (
            patch(
                "workload_client.Path.read_bytes",
                side_effect=[b"incomplete save", valid_content],
            ),
            patch("workload_client.sleep"),
        ):
            retried_items = read_workload_excel(path, "2026-0830")
        assert len(retried_items) == 2

        invalid_path = Path(directory) / "invalid.xlsx"
        invalid_path.write_bytes(b"not an xlsx workbook")
        try:
            with patch("workload_client.sleep"):
                read_workload_excel(invalid_path, "2026-0830")
        except Exception as exc:
            assert "不是有效的 .xlsx" in str(exc)
        else:
            raise AssertionError("无效的 xlsx 文件应被拒绝")

        context = WorkloadContext(
            plan_version="2026-0830",
            module="开发一组-前端",
            load_path="/task/load",
            save_path="/task/save",
            load_params={"planVersion": "2026-0830", "module": "开发一组-前端"},
            requirements={"SR001": "测试需求"},
            developers={"危国": Developer("危国", "weig", "前端")},
            existing_tasks=[],
            next_ord=3,
        )
        preview = build_workload_preview(items, context, ALL_GROUP_MEMBERS)
        assert len(preview.submittable_rows) == 1
        assert preview.error_count == 1
        assert preview.submittable_rows[0].record["developerId"] == "weig"
        assert preview.computed_hours == 16

        server_context = WorkloadContext(
            plan_version="2026-0830",
            module="开发一组",
            load_path="/task/getSvrDevTaskByPlanAndModule",
            save_path="/task/batchSaveServerDevTasks",
            load_params={"planVersion": "2026-0830", "module": "开发一组"},
            requirements={"SR001": "测试需求"},
            developers={"危国": Developer("危国", "weig", "后端")},
            existing_tasks=[],
            next_ord=1,
            ui_definition_path=SERVER_WORKLOAD_DEFINITION,
        )
        server_preview = build_workload_preview(
            [items[0]], server_context, ALL_GROUP_MEMBERS
        )
        assert len(server_preview.submittable_rows) == 1
        server_record = server_preview.submittable_rows[0].record
        assert server_record is not None
        assert server_record["methodType"] == "全新开发"
        assert server_record["reqAnalyzeWorkLoad"] == 2
        assert server_record["codeComplexity"] == 8
        assert server_record["unitTestNumber"] == 6
        assert server_preview.computed_hours == 16

        duplicate_record = dict(preview.submittable_rows[0].record)
        duplicate_context = WorkloadContext(
            plan_version=context.plan_version,
            module=context.module,
            load_path=context.load_path,
            save_path=context.save_path,
            load_params=context.load_params,
            requirements=context.requirements,
            developers=context.developers,
            existing_tasks=[duplicate_record],
            next_ord=4,
        )
        duplicate_preview = build_workload_preview(
            [items[0]], duplicate_context, ALL_GROUP_MEMBERS
        )
        assert duplicate_preview.duplicate_count == 1
        assert not duplicate_preview.submittable_rows


if __name__ == "__main__":
    test_analytics()
    test_crm_helpers()
    test_credential_store()
    test_all_detail_fields()
    test_crm_search_helpers()
    test_retrospective_markdown()
    test_workload_import()
    print("核心逻辑测试通过")
