"""缺陷数据清洗、筛选与汇总。"""

from __future__ import annotations

from collections import Counter
from typing import Any, Iterable


def text(value: Any) -> str:
    """将接口中的空值统一转换为空字符串。"""
    return "" if value is None else str(value).strip()


def introducer_name(bug: dict[str, Any]) -> str:
    """按绩效网站的规则取得“缺陷引入人”。"""
    return (
        text(bug.get("createdByName"))
        or text(bug.get("createdBy"))
        or text(bug.get("resolvedByName"))
        or text(bug.get("assignedToName"))
    )


def normalized_bug(bug: dict[str, Any]) -> dict[str, Any]:
    """补充供客户端展示的派生字段，不修改接口原始对象。"""
    result = dict(bug)
    result["_introducer"] = introducer_name(bug)
    return result


def filter_bugs(
    bugs: Iterable[dict[str, Any]],
    introducer: str = "",
    keyword: str = "",
) -> list[dict[str, Any]]:
    """按引入人精确筛选，并可按编号/标题/模块进行模糊搜索。"""
    wanted_introducer = text(introducer).casefold()
    wanted_keyword = text(keyword).casefold()
    result: list[dict[str, Any]] = []

    for raw_bug in bugs:
        bug = normalized_bug(raw_bug)
        if wanted_introducer and text(bug["_introducer"]).casefold() != wanted_introducer:
            continue
        if wanted_keyword:
            haystack = " ".join(
                [
                    text(bug.get("id")),
                    text(bug.get("title")),
                    text(bug.get("module")),
                    text(bug.get("status")),
                    text(bug.get("severity")),
                    text(bug.get("_introducer")),
                ]
            ).casefold()
            if wanted_keyword not in haystack:
                continue
        result.append(bug)
    return result


def count_by(items: Iterable[dict[str, Any]], field: str) -> Counter[str]:
    """按字段计数，空值显示为“未填写”"""
    return Counter(text(item.get(field)) or "未填写" for item in items)


def introducer_summary(bugs: Iterable[dict[str, Any]]) -> list[dict[str, Any]]:
    """生成按缺陷引入人汇总的数据。"""
    grouped: dict[str, list[dict[str, Any]]] = {}
    for raw_bug in bugs:
        bug = normalized_bug(raw_bug)
        grouped.setdefault(text(bug["_introducer"]) or "未填写", []).append(bug)

    rows: list[dict[str, Any]] = []
    for name, items in grouped.items():
        severities = count_by(items, "severity")
        statuses = count_by(items, "status")
        rows.append(
            {
                "introducer": name,
                "total": len(items),
                "severity_1": severities.get("1", 0),
                "severity_2": severities.get("2", 0),
                "severity_3": severities.get("3", 0),
                "severity_4": severities.get("4", 0),
                "closed": statuses.get("关闭", 0),
                "open": len(items) - statuses.get("关闭", 0),
            }
        )
    return sorted(rows, key=lambda item: (-item["total"], item["introducer"]))


def summary_metrics(bugs: Iterable[dict[str, Any]]) -> dict[str, int]:
    """生成顶部统计卡片需要的数据。"""
    items = list(bugs)
    severities = count_by(items, "severity")
    statuses = count_by(items, "status")
    return {
        "total": len(items),
        "severity_1": severities.get("1", 0),
        "severity_2": severities.get("2", 0),
        "severity_3": severities.get("3", 0),
        "severity_4": severities.get("4", 0),
        "closed": statuses.get("关闭", 0),
        "open": len(items) - statuses.get("关闭", 0),
    }
