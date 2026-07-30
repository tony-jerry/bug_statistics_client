"""根据绩效缺陷和 CRM 详情生成归纳型 Markdown Bug 回溯。"""

from __future__ import annotations

import re
from collections import Counter

from analytics import summary_metrics, text
from crm_client import BugDetail
from detail_utils import rich_text_to_plain


CATEGORY_RULES = (
    (
        "国际化多语言",
        ("国际化", "多语言", "翻译", "英文", "中文", "文案", "i18n"),
    ),
    (
        "页面交互与边界",
        ("选中态", "残留", "自适应", "弹窗", "边界", "交互", "缩放"),
    ),
    (
        "页面UI样式",
        ("样式", "对齐", "错位", "图标", "菜单", "列表", "布局", "展示"),
    ),
    (
        "数据与查询",
        ("查询", "搜索", "导出", "数据", "空值", "统计"),
    ),
    (
        "流程与权限",
        ("流程", "权限", "审批", "登录", "角色", "组织"),
    ),
)

ROOT_CAUSE_FALLBACKS = {
    "国际化多语言": (
        "多语言场景自测覆盖不足，边角弹窗、提示文案容易漏测，"
        "中英文环境适配不完善"
    ),
    "页面交互与边界": (
        "页面边界交互考虑不周，选中态、弹窗及自适应布局等特殊场景"
        "验证不充分"
    ),
    "页面UI样式": (
        "页面UI样式校验不严，菜单、列表、弹窗和详情页存在图标或"
        "内容对齐问题"
    ),
    "数据与查询": (
        "查询与数据边界覆盖不足，对空值、多条件及导出场景缺少完整验证"
    ),
    "流程与权限": (
        "流程及权限组合场景覆盖不足，对不同角色和组织范围缺少交叉验证"
    ),
}

IMPROVEMENTS = {
    "国际化多语言": (
        "补齐多语言自测用例，所有迭代需求必须完成中英文双环境回归，"
        "杜绝文案残留、翻译缺失问题"
    ),
    "页面交互与边界": (
        "完善页面边界、交互残留、自适应布局等边界场景自测，覆盖"
        "搜索选中、弹窗提示、页面缩放等特殊场景"
    ),
    "页面UI样式": (
        "统一前端UI布局规范，针对列表、菜单、弹窗、详情页做统一"
        "样式兜底，降低同类展示问题复发率"
    ),
    "数据与查询": (
        "补充查询、空值、多条件组合及导出场景回归，完善数据异常时"
        "的提示与兜底验证"
    ),
    "流程与权限": (
        "补充不同角色、组织及流程节点的权限矩阵测试，覆盖关键操作"
        "的正向与反向场景"
    ),
}


def retrospective_version_label(plan_version: str) -> str:
    value = plan_version.strip()
    if "-" in value:
        value = value.rsplit("-", 1)[-1]
    return value or "未命名"


def retrospective_filename(plan_version: str, introducer: str) -> str:
    version = retrospective_version_label(plan_version)
    reporter = introducer.strip() or "全部"
    for character in '<>:"/\\|?*':
        reporter = reporter.replace(character, "_")
    return f"{version}版本bug回溯-{reporter}.md"


def _clean_plain_text(value: str) -> str:
    plain = rich_text_to_plain(value).replace("[图片]", " ")
    plain = re.sub(r"https?://\S+", " ", plain)
    plain = re.sub(r"【[^】]*(?:环境|信息)[^】]*】", " ", plain)
    plain = re.sub(r"\s+", " ", plain)
    return plain.strip(" ：:，,；;。")


def _field_text(
    detail: BugDetail | None,
    keys: tuple[str, ...],
    labels: tuple[str, ...] = (),
    fallback: str = "",
) -> str:
    if detail:
        key_set = {item.lower() for item in keys}
        for field in detail.fields:
            if (
                field.key.lower() in key_set
                or any(label in field.label for label in labels)
            ):
                value = _clean_plain_text(field.value)
                if value:
                    return value
    return _clean_plain_text(fallback)


def _detail_description(detail: BugDetail | None) -> str:
    return _field_text(
        detail,
        ("description",),
        ("问题描述", "描述内容", "描述"),
        detail.description_html if detail else "",
    )


def _detail_root_cause(detail: BugDetail | None) -> str:
    return _field_text(
        detail,
        ("dts_rootcauseanalysis",),
        ("根因分析", "产生原因", "原因分析"),
        detail.root_cause_html if detail else "",
    )


def _detail_fix_content(detail: BugDetail | None) -> str:
    return _field_text(
        detail,
        ("dts_fixcontent",),
        ("修复内容", "解决方案", "整改内容"),
        detail.fix_content_html if detail else "",
    )


def _bug_search_text(bug: dict, detail: BugDetail | None) -> str:
    values = [
        text(bug.get("title")),
        text(bug.get("module")),
        _detail_description(detail),
        _detail_root_cause(detail),
        _detail_fix_content(detail),
    ]
    if detail:
        values.extend(_clean_plain_text(field.value) for field in detail.fields)
    return " ".join(values).lower()


def _classify_category(search_text: str) -> str:
    scores = {
        name: sum(search_text.count(keyword.lower()) for keyword in keywords)
        for name, keywords in CATEGORY_RULES
    }
    category, score = max(scores.items(), key=lambda item: item[1])
    return category if score else "其他问题"


def _invalid_reason(search_text: str) -> str:
    rules = (
        ("无法复现", ("无法复现", "不能复现", "未复现")),
        ("重复问题单", ("重复问题", "重复缺陷", "重复单", "重复提单")),
        ("非前端实际问题", ("非前端", "后端问题", "非本端")),
        ("非缺陷", ("非缺陷", "不是缺陷", "非问题", "无效问题")),
    )
    for name, keywords in rules:
        if any(keyword in search_text for keyword in keywords):
            return name
    return ""


def _shorten(value: str, limit: int = 150) -> str:
    cleaned = value.strip()
    if len(cleaned) <= limit:
        return cleaned
    shortened = cleaned[:limit].rsplit("，", 1)[0].strip()
    return (shortened or cleaned[:limit]).rstrip("。") + "…"


def _strip_title_prefixes(value: str) -> str:
    cleaned = re.sub(r"^\s*(?:【[^】]+】|\[[^\]]+\])+\s*", "", value)
    return cleaned.strip() or value.strip()


def _invalid_breakdown(reasons: Counter[str]) -> str:
    if not reasons:
        return ""
    return "（" + "、".join(
        f"{count}条{reason}" for reason, count in reasons.items()
    ) + "）"


def _module_summary(bugs: list[dict]) -> str:
    modules: list[str] = []
    for bug in bugs:
        module = text(bug.get("module")).strip()
        if module and module not in modules:
            modules.append(module)
    return "、".join(modules) if modules else "未标注模块"


def build_retrospective_markdown(
    plan_version: str,
    introducer: str,
    bugs: list[dict],
    details: dict[str, BugDetail | None],
) -> str:
    version = retrospective_version_label(plan_version)
    reporter = introducer.strip() or "全部"
    metrics = summary_metrics(bugs)

    records: list[dict[str, object]] = []
    invalid_reasons: Counter[str] = Counter()
    category_counts: Counter[str] = Counter()
    collaborator_names: list[str] = []
    root_causes: list[str] = []

    for bug in bugs:
        bug_number = text(bug.get("id"))
        detail = details.get(bug_number)
        search_text = _bug_search_text(bug, detail)
        invalid_reason = _invalid_reason(search_text)
        category = _classify_category(search_text)
        if invalid_reason:
            invalid_reasons[invalid_reason] += 1
        else:
            category_counts[category] += 1

        resolver = text(bug.get("resolvedByName")).strip()
        if resolver and reporter not in resolver and resolver not in collaborator_names:
            collaborator_names.append(resolver)

        root_cause = _detail_root_cause(detail)
        if root_cause and root_cause not in root_causes:
            root_causes.append(root_cause)

        records.append(
            {
                "bug": bug,
                "detail": detail,
                "category": category,
                "invalid": bool(invalid_reason),
                "root": root_cause,
                "fix": _detail_fix_content(detail),
                "description": _detail_description(detail),
            }
        )

    invalid_count = sum(invalid_reasons.values())
    valid_count = max(0, len(bugs) - invalid_count)
    collaboration_count = sum(
        1
        for record in records
        if (
            text(record["bug"].get("resolvedByName")).strip()
            and reporter not in text(record["bug"].get("resolvedByName"))
        )
    )
    self_fixed_count = max(0, valid_count - collaboration_count)
    primary_category, primary_count = (
        category_counts.most_common(1)[0]
        if category_counts
        else ("其他问题", 0)
    )

    lines = [
        f"# {version}版本 bug回溯",
        "",
        f"**周期**：{version}版本SIT测试",
        "",
        f"**模块**：{_module_summary(bugs)}",
        "",
        f"**汇报人**：{reporter}",
        "",
        "## 一、缺陷概况",
        "",
        f"- 总处理缺陷：**{metrics['total']}条**",
        "",
        (
            f"- 无效缺陷：{invalid_count}条"
            f"{_invalid_breakdown(invalid_reasons)}"
        ),
        "",
        f"- 自主有效修复缺陷：{self_fixed_count}条",
        "",
    ]
    if primary_count:
        lines.extend(
            [
                (
                    f"- {primary_category}类缺陷：**{primary_count}条**，"
                    "为本版本主要问题类型"
                ),
                "",
            ]
        )
    if collaboration_count:
        collaborator_text = "、".join(collaborator_names)
        lines.extend(
            [
                (
                    f"- 协同处理：**{collaboration_count}条协同bug修复**"
                    + (f"（涉及{collaborator_text}）" if collaborator_text else "")
                ),
                "",
            ]
        )
    lines.extend(
        [
            (
                f"- 状态：已关闭 {metrics['closed']} 条，"
                f"未关闭 {metrics['open']} 条；有效问题均按当前状态持续闭环"
            ),
            "",
            "## 二、主要问题根因",
            "",
        ]
    )

    summarized_roots = [_shorten(item) for item in root_causes[:5]]
    used_roots = set(summarized_roots)
    for category, _count in category_counts.most_common():
        fallback = ROOT_CAUSE_FALLBACKS.get(category)
        if fallback and fallback not in used_roots and len(summarized_roots) < 5:
            summarized_roots.append(fallback)
            used_roots.add(fallback)
    if not summarized_roots:
        summarized_roots.append(
            "自测与回归场景覆盖不足，对边界条件和异常流程验证不充分"
        )
    for index, root in enumerate(summarized_roots, 1):
        prefix = "核心诱因：" if index == 1 else ""
        lines.extend([f"{index}. {prefix}{root}", ""])

    candidates = [
        record
        for record in records
        if not record["invalid"] and record["category"] == primary_category
    ] or [record for record in records if not record["invalid"]] or records
    typical = max(
        candidates,
        key=lambda record: len(str(record["root"])) + len(str(record["fix"])),
        default=None,
    )
    if typical:
        typical_bug = typical["bug"]
        problem = _strip_title_prefixes(text(typical_bug.get("title")))
        if not problem:
            problem = _shorten(str(typical["description"])) or "（未填写）"
        root = _shorten(str(typical["root"])) or ROOT_CAUSE_FALLBACKS.get(
            str(typical["category"]),
            "相关场景自测覆盖不足，问题识别和验证不充分",
        )
        fixed = _shorten(str(typical["fix"])) or "已完成问题整改并补充回归验证"
    else:
        problem = "（当前筛选结果无可分析缺陷）"
        root = "（未填写）"
        fixed = "（未填写）"

    lines.extend(
        [
            "## 三、典型Bug深度分析",
            "",
            "**典型Bug深度分析**",
            "",
            f"**问题**：{problem}",
            "",
            f"**根因**：{root}",
            "",
            f"**已修复**：{fixed}",
            "",
            "## 四、后续改进",
            "",
        ]
    )

    improvement_lines: list[str] = []
    for category, _count in category_counts.most_common():
        improvement = IMPROVEMENTS.get(category)
        if improvement and improvement not in improvement_lines:
            improvement_lines.append(improvement)
    if invalid_reasons.get("重复问题单"):
        improvement_lines.append(
            "提单前提前检索历史问题单，主动筛查重复bug，减少无效问题单流转"
        )
    if len(improvement_lines) < 3:
        improvement_lines.extend(
            [
                "将本版本问题沉淀为回归用例，并在后续版本发布前完成复测",
                "对根因分析和修复内容未填写的缺陷及时补齐记录，确保问题可追溯",
            ]
        )
    for index, improvement in enumerate(dict.fromkeys(improvement_lines[:4]), 1):
        lines.extend([f"{index}. {improvement}", ""])

    lines.extend(["> （注：部分内容可能由 AI 生成）", ""])
    return "\n".join(lines)
