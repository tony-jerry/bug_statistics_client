"""CRM 登录、流程定位与缺陷详情读取。"""

from __future__ import annotations

import base64
import binascii
import json
import re
from dataclasses import dataclass
from typing import Any
from urllib.parse import unquote, urlencode, urljoin, urlparse

import requests
from cryptography.hazmat.primitives import padding
from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes


DEFAULT_CRM_BASE_URL = "https://crm.hustcad.com"
CRM_SEARCH_PAGE_SIZE = 50
CRM_SEARCH_COLUMNS = (
    ("objectNumber", "编号", 180),
    ("title", "标题", 360),
    ("defectCategory", "缺陷类别", 120),
    ("827308980874936320", "迭代版本", 120),
    ("lifecyclestagekeyName", "状态", 90),
    ("773662542538473472", "需求模块", 180),
    ("severity", "严重程度", 90),
    ("priorityName", "优先级", 90),
    ("leader", "负责人", 130),
    ("solution", "解决方案", 120),
    ("creatorDisplayName", "创建者", 130),
    ("createstamp", "创建时间", 150),
    ("modifystamp", "最后修改时间", 150),
)
_AES_KEY_SOURCE = "hYwy5Fs0neRUuM1Pf+/NjQ=="
_AES_IV = b"1234567890abcdef"
_LEGACY_FIELD_LABELS = {
    "severitykey": {
        "fatal": "致命",
        "blocker": "致命",
        "critical": "严重",
        "major": "一般",
        "minor": "提示",
    },
    "categorykey": {
        "codeerror": "代码错误",
        "performance": "性能问题",
        "design": "设计缺陷",
        "ui": "UI界面问题",
        "other": "其他",
    },
    "causekey": {
        "codingerror": "编码错误",
    },
    "solutionkey": {
        "fixed": "已修复",
    },
}


class CrmClientError(RuntimeError):
    """可直接展示给用户的 CRM 错误。"""


class CrmAuthenticationError(CrmClientError):
    """CRM 登录或令牌错误。"""


@dataclass(frozen=True)
class CrmLoginResult:
    username: str


@dataclass(frozen=True)
class BugDetailField:
    key: str
    label: str
    value: str
    is_rich_text: bool = False


@dataclass(frozen=True)
class BugDetail:
    workflow_key: str
    object_oid: str
    bug_number: str
    title: str
    description_html: str
    root_cause_html: str
    fix_content_html: str
    fields: tuple[BugDetailField, ...]


@dataclass(frozen=True)
class CrmBugPage:
    items: tuple[dict[str, str], ...]
    total: int
    page_num: int
    page_size: int
    pages: int
    keyword: str


def encrypt_crm_password(password: str) -> str:
    """复刻 CRM 网页端 CryptoJS AES-CBC + 双层 Base64 加密。"""
    key = base64.b64encode(_AES_KEY_SOURCE.encode("utf-8"))
    padder = padding.PKCS7(128).padder()
    padded = padder.update(password.encode("utf-8")) + padder.finalize()
    encryptor = Cipher(algorithms.AES(key), modes.CBC(_AES_IV)).encryptor()
    ciphertext = encryptor.update(padded) + encryptor.finalize()
    inner_base64 = base64.b64encode(ciphertext)
    return base64.b64encode(inner_base64).decode("ascii")


def extract_crm_object_identity(bug_url: str) -> tuple[str, str]:
    """从绩效接口返回的多层 URL 编码参数中提取 CRM 对象 oid/otype。"""
    decoded = bug_url or ""
    for _ in range(6):
        next_value = unquote(decoded)
        if next_value == decoded:
            break
        decoded = next_value

    oid_match = re.search(r'"oid"\s*:\s*"(\d+)"', decoded)
    otype_match = re.search(r'"otype"\s*:\s*"([^"]+)"', decoded)
    if not oid_match or not otype_match:
        raise CrmClientError("缺陷链接中未找到 CRM 对象标识")
    return oid_match.group(1), otype_match.group(1)


def build_crm_bug_url(
    base_url: str,
    object_oid: str,
    object_otype: str,
    label: str = "",
) -> str:
    """构造可在浏览器中直接打开的 CRM 缺陷详情地址。"""
    query = urlencode(
        {
            "otype": object_otype,
            "oid": object_oid,
            "label": label,
        }
    )
    return (
        f"{base_url.rstrip('/')}/#/homePage/defectObjForm?"
        f"{query}"
    )


def _join_rich_values(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, list):
        return "\n<hr>\n".join(str(item) for item in value if item)
    return str(value)


def _json_object(value: Any) -> dict[str, Any]:
    if isinstance(value, dict):
        return value
    if isinstance(value, str) and value.strip():
        try:
            loaded = json.loads(value)
        except ValueError:
            return {}
        if isinstance(loaded, dict):
            return loaded
    return {}


def _option_labels(options: dict[str, Any]) -> dict[str, str]:
    labels: dict[str, str] = {}
    items = options.get("optionItems") or []
    if not isinstance(items, list):
        return labels
    for item in items:
        if not isinstance(item, dict):
            continue
        label = str(
            item.get("label")
            or item.get("displayName")
            or item.get("description")
            or item.get("value")
            or ""
        ).strip()
        if not label:
            continue
        for candidate in (
            item.get("value"),
            item.get("internalName"),
            item.get("name"),
            item.get("oid"),
        ):
            if candidate is not None:
                labels[str(candidate)] = label
                labels[str(candidate).lower()] = label
    return labels


def _format_field_value(
    value: Any,
    labels: dict[str, str],
    field_key: str = "",
) -> str:
    if value is None or value == "" or value == [] or value == {}:
        return "（空）"
    if isinstance(value, bool):
        return "是" if value else "否"
    if isinstance(value, dict):
        for key in (
            "displayFullName",
            "displayName",
            "fullname",
            "fullName",
            "label",
            "name",
            "objectnumber",
            "number",
            "username",
            "userName",
            "account",
        ):
            display = value.get(key)
            if display not in (None, ""):
                return str(display)
        return json.dumps(value, ensure_ascii=False, sort_keys=True)
    if isinstance(value, (list, tuple)):
        formatted = [
            _format_field_value(item, labels, field_key)
            for item in value
            if item not in (None, "")
        ]
        return "、".join(item for item in formatted if item != "（空）") or "（空）"
    raw = str(value)
    legacy_labels = _LEGACY_FIELD_LABELS.get(field_key, {})
    return labels.get(
        raw,
        labels.get(raw.lower(), legacy_labels.get(raw.lower(), raw)),
    )


def extract_detail_fields(
    form_data: dict[str, Any],
    form_template: dict[str, Any],
) -> tuple[BugDetailField, ...]:
    """按 CRM 表单顺序提取所有业务字段，包括未填写字段。"""
    iba = form_data.get("ibaAttribute") or {}
    if not isinstance(iba, dict):
        iba = {}

    fields: list[BugDetailField] = []
    seen: set[str] = set()
    seen_labels: set[str] = set()
    widgets = form_template.get("widgetList") or []
    if isinstance(widgets, list):
        for widget in widgets:
            if not isinstance(widget, dict):
                continue
            options = widget.get("options") or {}
            if not isinstance(options, dict):
                continue
            key = str(options.get("realProps") or "").strip()
            label = str(options.get("label") or "").strip()
            if not key or not label or key in seen:
                continue
            raw_value = form_data.get(key) if key in form_data else iba.get(key)
            is_rich_text = str(widget.get("type") or "") == "rich-editor"
            value = (
                _join_rich_values(raw_value)
                if is_rich_text
                else _format_field_value(
                    raw_value,
                    _option_labels(options),
                    key,
                )
            )
            fields.append(
                BugDetailField(
                    key=key,
                    label=label,
                    value=value,
                    is_rich_text=is_rich_text,
                )
            )
            seen.add(key)
            seen_labels.add(label)

    attributes = form_template.get("attributeViewDetailList") or []
    if isinstance(attributes, list):
        for item in attributes:
            if not isinstance(item, dict):
                continue
            attribute = item.get("attribute") or {}
            if not isinstance(attribute, dict):
                continue
            key = str(attribute.get("name") or "").strip()
            label = str(attribute.get("displayName") or key).strip()
            if not key or key in seen or label in seen_labels:
                continue
            raw_value = form_data.get(key) if key in form_data else iba.get(key)
            data_type = str(attribute.get("dataTypeName") or "")
            is_rich_text = (
                "HTMLText" in data_type
                or key
                in {
                    "DTS_RootCauseAnalysis",
                    "DTS_FixContent",
                    "DTS_TestResult",
                }
            )
            value = (
                _join_rich_values(raw_value)
                if is_rich_text
                else _format_field_value(raw_value, {}, key)
            )
            fields.append(
                BugDetailField(
                    key=key,
                    label=label,
                    value=value,
                    is_rich_text=is_rich_text,
                )
            )
            seen.add(key)
            seen_labels.add(label)
    return tuple(fields)


def _crm_search_column_payloads() -> list[dict[str, Any]]:
    property_types = {
        "objectNumber": "String",
        "title": "String",
        "defectCategory": "String",
        "severity": "String",
        "priorityName": "String",
        "leader": "String",
        "solution": "String",
        "createstamp": "Time",
        "modifystamp": "Time",
    }
    iba_types = {
        "827308980874936320": "ty.inteplm.attribute.CTyStringDef",
        "773662542538473472": "ty.inteplm.attribute.CTyStringDef",
    }
    columns: list[dict[str, Any]] = []
    for property_name, display_name, _width in CRM_SEARCH_COLUMNS:
        column: dict[str, Any] = {
            "propertyName": property_name,
            "displayName": display_name,
        }
        if property_name in property_types:
            column["propertyType"] = property_types[property_name]
        if property_name in iba_types:
            column["ibaDefOid"] = property_name
            column["ibaDefOtype"] = iba_types[property_name]
        columns.append(column)
    return columns


def build_crm_search_payload(
    keyword: str,
    page_num: int,
    *,
    need_count: bool,
) -> dict[str, Any]:
    """构造 CRM 缺陷列表或总数请求；列表页码从 1 开始。"""
    type_short = {
        "typeOid": "598591129061539840",
        "typeInthid": "ty.inteplm.ipd.CTyDefect",
        "displayName": "缺陷",
    }
    type_detail = {
        **type_short,
        "otype": "ty.inteplm.type.CTyTypeDef",
        "name": "ty.inteplm.ipd.CTyDefect",
        "instantiable": "1",
        "logicalidentifier": "ty.inteplm.ipd.CTyDefect",
        "foreignKeyObjectDTO": {
            "relationship": "master",
            "fieldName": "projectoid",
            "fieldValue": "770728653012500480",
        },
    }
    columns = _crm_search_column_payloads()
    requested_page = 0 if need_count else max(1, int(page_num))
    condition = {
        "propertyType": "String",
        "propertyName": "number",
        "typeList": [dict(type_short)],
        "multivalued": False,
        "ignoreCase": True,
        "symbol": "!=",
        "propertyValue": "0",
        "empty": True,
        "classification": False,
        "iba": False,
    }
    view_condition = {
        "typeList": [dict(type_short)],
        "contextList": [],
        "conditionJoint": "and",
        "conditionGroups": [
            {
                "conditionJoint": "or",
                "conditions": [condition],
            }
        ],
        "keyword": "",
        "needChildrenClassification": False,
        "pageNum": max(1, requested_page),
        "pageSize": CRM_SEARCH_PAGE_SIZE,
        "needIBA": False,
        "columns": [dict(column) for column in columns],
        "needThumbnailFile": False,
        "needParticipateContainer": False,
        "searchType": 0,
    }
    payload: dict[str, Any] = {
        "needSort": True,
        "typeList": [type_detail],
        "keyword": keyword.strip(),
        "columns": columns,
        "pageNum": requested_page,
        "pageSize": CRM_SEARCH_PAGE_SIZE,
        "sorts": {},
        "viewCondition": view_condition,
        "conditionJoint": "and",
        "needCount": need_count,
    }
    if not need_count:
        payload["isSetColumn"] = False
    return payload


def normalize_crm_bug(item: dict[str, Any]) -> dict[str, str]:
    """把 CRM 搜索结果中的扩展字段整理成表格可直接展示的数据。"""
    extension = item.get("extAttrMapForSearch") or {}
    if not isinstance(extension, dict):
        extension = {}
    iba = item.get("ibaAttrMapForSearch") or {}
    if not isinstance(iba, dict):
        iba = {}

    def value(key: str, *fallbacks: str) -> str:
        candidates = [item.get(key), extension.get(key), iba.get(key)]
        candidates.extend(item.get(name) for name in fallbacks)
        for candidate in candidates:
            if candidate not in (None, ""):
                return str(candidate)
        return ""

    return {
        "objectNumber": value("objectNumber"),
        "title": value("title", "name", "displayFullName"),
        "defectCategory": value("defectCategory"),
        "827308980874936320": value("827308980874936320"),
        "lifecyclestagekeyName": value("lifecyclestagekeyName"),
        "773662542538473472": value("773662542538473472"),
        "severity": value("severity"),
        "priorityName": value("priorityName"),
        "leader": value("leader"),
        "solution": value("solution"),
        "creatorDisplayName": value("creatorDisplayName", "creatorFullName"),
        "createstamp": value("createstamp"),
        "modifystamp": value("modifystamp"),
        "_oid": value("realOid", "oid"),
        "_otype": value("realOtype", "otype"),
    }


class CrmApiClient:
    """CRM 会话仅驻留内存，不持久化密码或令牌。"""

    def __init__(
        self,
        base_url: str = DEFAULT_CRM_BASE_URL,
        timeout: int = 30,
    ) -> None:
        self.base_url = base_url.rstrip("/")
        self.timeout = timeout
        self.session = requests.Session()
        self.username = ""
        self._workflow_key_cache: dict[str, str] = {}
        self._detail_cache: dict[str, BugDetail] = {}
        self._validate_base_url()

    def _validate_base_url(self) -> None:
        parsed = urlparse(self.base_url)
        if parsed.scheme not in {"http", "https"} or not parsed.netloc:
            raise ValueError("CRM 地址必须是有效的 http/https 地址")

    def _url(self, path: str) -> str:
        return urljoin(f"{self.base_url}/", path.lstrip("/"))

    @staticmethod
    def _json(response: requests.Response, action: str) -> dict[str, Any]:
        try:
            payload = response.json()
        except ValueError as exc:
            raise CrmClientError(f"{action}返回了无法识别的数据") from exc
        if not isinstance(payload, dict):
            raise CrmClientError(f"{action}返回的数据格式异常")
        return payload

    @staticmethod
    def _message(payload: dict[str, Any], fallback: str) -> str:
        return str(
            payload.get("message")
            or payload.get("msg")
            or (payload.get("errors") or {}).get("message")
            or fallback
        )

    def login(self, username: str, password: str) -> CrmLoginResult:
        username = username.strip()
        if not username:
            raise CrmAuthenticationError("请输入 CRM 用户名")
        if not password:
            raise CrmAuthenticationError("请输入 CRM 密码")

        self.session.close()
        self.session = requests.Session()
        try:
            response = self.session.post(
                self._url("/rest/userService/v1/user/userLoginPlm"),
                json={
                    "name": username,
                    "password": encrypt_crm_password(password),
                    "appID": "Chrome(150.0.0.0)",
                    "passwordFlag": "1",
                },
                timeout=self.timeout,
            )
        except requests.RequestException as exc:
            raise CrmAuthenticationError(f"CRM 登录请求失败：{exc}") from exc

        payload = self._json(response, "CRM 登录接口")
        authorization = response.headers.get("authorization")
        if response.status_code != 200 or not payload.get("success") or not authorization:
            raise CrmAuthenticationError(self._message(payload, "CRM 登录失败"))

        self.session.headers.update({"Authorization": authorization})
        self.username = username
        self._workflow_key_cache.clear()
        self._detail_cache.clear()
        return CrmLoginResult(username=username)

    def _post_json(self, path: str, data: dict[str, Any], action: str) -> dict[str, Any]:
        if "Authorization" not in self.session.headers:
            raise CrmAuthenticationError("请先登录 CRM")
        try:
            response = self.session.post(
                self._url(path),
                json=data,
                timeout=self.timeout,
            )
        except requests.RequestException as exc:
            raise CrmClientError(f"{action}请求失败：{exc}") from exc
        if response.status_code in {401, 403}:
            raise CrmAuthenticationError("CRM 登录状态已失效，请重新登录")
        payload = self._json(response, action)
        if response.status_code != 200 or not payload.get("success"):
            raise CrmClientError(self._message(payload, f"{action}失败"))
        return payload

    def resolve_workflow_key(self, bug_url: str) -> tuple[str, str]:
        object_oid, object_otype = extract_crm_object_identity(bug_url)
        cached = self._workflow_key_cache.get(object_oid)
        if cached:
            return object_oid, cached

        payload = self._post_json(
            "/rest/v1/formData/getWorkflowContainerIdByOid",
            {"oid": object_oid, "otype": object_otype},
            "查询缺陷流程",
        )
        data = payload.get("data") or {}
        workflow_key = str(data.get("workflowContainerOid") or "").strip()
        if not workflow_key:
            raise CrmClientError("该缺陷没有可读取的流程详情")
        self._workflow_key_cache[object_oid] = workflow_key
        return object_oid, workflow_key

    def search_bugs(self, keyword: str, page_num: int = 1) -> CrmBugPage:
        count_payload = self._post_json(
            "/rest/v1/search/queryByConditionForPage",
            build_crm_search_payload(keyword, 0, need_count=True),
            "查询 CRM 缺陷总数",
        )
        count_data = count_payload.get("data") or {}
        if not isinstance(count_data, dict):
            raise CrmClientError("CRM 缺陷总数返回格式异常")
        try:
            total = max(0, int(count_data.get("total") or 0))
        except (TypeError, ValueError) as exc:
            raise CrmClientError("CRM 缺陷总数无法识别") from exc

        pages = (
            (total + CRM_SEARCH_PAGE_SIZE - 1) // CRM_SEARCH_PAGE_SIZE
            if total
            else 0
        )
        actual_page = max(1, int(page_num))
        if pages:
            actual_page = min(actual_page, pages)
        list_payload = self._post_json(
            "/rest/v1/search/queryByConditionForPage",
            build_crm_search_payload(
                keyword,
                actual_page,
                need_count=False,
            ),
            "查询 CRM 缺陷列表",
        )
        list_data = list_payload.get("data") or {}
        if not isinstance(list_data, dict):
            raise CrmClientError("CRM 缺陷列表返回格式异常")
        raw_items = list_data.get("list") or []
        if not isinstance(raw_items, list):
            raise CrmClientError("CRM 缺陷列表数据格式异常")
        items = tuple(
            normalize_crm_bug(item)
            for item in raw_items
            if isinstance(item, dict)
        )
        return CrmBugPage(
            items=items,
            total=total,
            page_num=actual_page,
            page_size=CRM_SEARCH_PAGE_SIZE,
            pages=pages,
            keyword=keyword.strip(),
        )

    def get_bug_detail_by_identity(
        self,
        object_oid: str,
        object_otype: str,
    ) -> BugDetail:
        identity = json.dumps(
            {
                "oid": object_oid,
                "otype": object_otype,
            }
        )
        return self.get_bug_detail(identity)

    def get_bug_detail(self, bug_url: str) -> BugDetail:
        object_oid, workflow_key = self.resolve_workflow_key(bug_url)
        if workflow_key in self._detail_cache:
            return self._detail_cache[workflow_key]

        try:
            response = self.session.get(
                self._url(
                    "/rest/v1/workFlowContainer/"
                    "getWorkFlowContainerDetailByPrimaryKey"
                ),
                params={"key": workflow_key},
                timeout=self.timeout,
            )
        except requests.RequestException as exc:
            raise CrmClientError(f"读取缺陷详情失败：{exc}") from exc
        if response.status_code in {401, 403}:
            raise CrmAuthenticationError("CRM 登录状态已失效，请重新登录")

        payload = self._json(response, "缺陷详情接口")
        if response.status_code != 200 or not payload.get("success"):
            raise CrmClientError(self._message(payload, "读取缺陷详情失败"))

        data = payload.get("data") or {}
        structured = data.get("structuredFormDataDetailVO") or {}
        if isinstance(structured, str):
            try:
                structured = json.loads(structured)
            except ValueError as exc:
                raise CrmClientError("缺陷详情中的表单数据无法解析") from exc
        if not isinstance(structured, dict):
            raise CrmClientError("缺陷详情中的表单数据格式异常")

        form_data = structured.get("formData") or {}
        if not isinstance(form_data, dict):
            raise CrmClientError("缺陷详情中未找到表单内容")
        iba = form_data.get("ibaAttribute") or {}
        if not isinstance(iba, dict):
            iba = {}
        form_template = _json_object(structured.get("formtemplateData"))
        fields = extract_detail_fields(form_data, form_template)

        detail = BugDetail(
            workflow_key=workflow_key,
            object_oid=object_oid,
            bug_number=str(
                form_data.get("objectnumber")
                or form_data.get("originalObjectnumber")
                or ""
            ),
            title=str(form_data.get("name") or data.get("name") or ""),
            description_html=_join_rich_values(form_data.get("description")),
            root_cause_html=_join_rich_values(
                iba.get("DTS_RootCauseAnalysis")
            ),
            fix_content_html=_join_rich_values(iba.get("DTS_FixContent")),
            fields=fields,
        )
        self._detail_cache[workflow_key] = detail
        return detail

    def download_rich_images(self, sources: list[str]) -> dict[str, bytes]:
        """使用当前 CRM 令牌下载富文本图片，只允许访问 CRM 自身地址。"""
        if "Authorization" not in self.session.headers:
            raise CrmAuthenticationError("请先登录 CRM")

        images: dict[str, bytes] = {}
        base = urlparse(self.base_url)
        for source in dict.fromkeys(item.strip() for item in sources if item.strip()):
            if source.startswith("data:image/"):
                try:
                    _metadata, encoded = source.split(",", 1)
                    images[source] = base64.b64decode(encoded)
                except (ValueError, binascii.Error):
                    continue
                continue

            image_url = urljoin(f"{self.base_url}/", source)
            parsed = urlparse(image_url)
            if (
                parsed.scheme not in {"http", "https"}
                or parsed.netloc.lower() != base.netloc.lower()
            ):
                continue
            try:
                response = self.session.get(image_url, timeout=self.timeout)
            except requests.RequestException:
                continue
            if response.status_code in {401, 403}:
                raise CrmAuthenticationError("CRM 登录状态已失效，请重新登录")
            if response.status_code != 200:
                continue
            if len(response.content) > 20 * 1024 * 1024:
                continue
            images[source] = response.content
        return images

    def close(self) -> None:
        self.session.close()
        self.username = ""
        self._workflow_key_cache.clear()
        self._detail_cache.clear()
