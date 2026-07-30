"""绩效系统 SSO 登录与缺陷接口客户端。"""

from __future__ import annotations

from dataclasses import dataclass
from html.parser import HTMLParser
from typing import Any
from urllib.parse import urljoin, urlparse

import requests


DEFAULT_BASE_URL = "http://10.8.190.76:9000"


class BugClientError(RuntimeError):
    """可直接展示给用户的客户端错误。"""


class AuthenticationError(BugClientError):
    """登录或会话失效。"""


class _LoginFormParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.action: str | None = None

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag.lower() == "form" and self.action is None:
            self.action = dict(attrs).get("action")


@dataclass(frozen=True)
class LoginResult:
    display_name: str
    staff_number: str
    role: str


@dataclass(frozen=True)
class PlanVersion:
    year: int
    month: int
    plan_version: str


class BugApiClient:
    """持有内存会话；密码和令牌均不会写入磁盘。"""

    def __init__(self, base_url: str = DEFAULT_BASE_URL, timeout: int = 30) -> None:
        self.base_url = base_url.rstrip("/")
        self.timeout = timeout
        self.session = requests.Session()
        self.user_info: dict[str, Any] = {}
        self._validate_base_url()

    def _validate_base_url(self) -> None:
        parsed = urlparse(self.base_url)
        if parsed.scheme not in {"http", "https"} or not parsed.netloc:
            raise ValueError("服务地址必须是有效的 http/https 地址")

    def _url(self, path: str) -> str:
        return urljoin(f"{self.base_url}/", path.lstrip("/"))

    def login(self, username: str, password: str) -> LoginResult:
        username = username.strip()
        if not username:
            raise AuthenticationError("请输入用户名")
        if not password:
            raise AuthenticationError("请输入密码")

        self.session.close()
        self.session = requests.Session()

        try:
            auth_page = self.session.get(
                self._url("/rest/login/sso/authorization/inteplm"),
                timeout=self.timeout,
            )
            auth_page.raise_for_status()
        except requests.RequestException as exc:
            raise AuthenticationError(f"无法打开统一认证页面：{exc}") from exc

        parser = _LoginFormParser()
        parser.feed(auth_page.text)
        if not parser.action:
            raise AuthenticationError("统一认证页面中未找到登录表单")

        form_url = urljoin(auth_page.url, parser.action)
        try:
            login_response = self.session.post(
                form_url,
                data={
                    "username": username,
                    "password": password,
                    "credentialId": "",
                },
                allow_redirects=True,
                timeout=self.timeout,
            )
            login_response.raise_for_status()
        except requests.RequestException as exc:
            raise AuthenticationError(f"统一身份认证请求失败：{exc}") from exc

        if "login-actions/authenticate" in login_response.url:
            raise AuthenticationError("用户名或密码错误")
        if "sso=success" not in login_response.url:
            raise AuthenticationError("统一身份认证未完成，请确认账号状态")

        try:
            exchange = self.session.post(
                self._url("/rest/login/sso/exchange"),
                timeout=self.timeout,
            )
        except requests.RequestException as exc:
            raise AuthenticationError(f"登录会话交换失败：{exc}") from exc

        if exchange.status_code == 401:
            raise AuthenticationError("登录会话已失效，请重新登录")
        try:
            payload = exchange.json()
        except ValueError as exc:
            raise AuthenticationError("登录接口返回了无法识别的数据") from exc

        authorization = exchange.headers.get("authorization")
        if exchange.status_code != 200 or payload.get("result") != "SUCCESS" or not authorization:
            message = payload.get("message") or payload.get("msg") or "登录失败"
            raise AuthenticationError(str(message))

        self.session.headers.update({"Authorization": authorization})
        self.user_info = payload.get("data") or {}
        return LoginResult(
            display_name=str(self.user_info.get("name") or username),
            staff_number=str(self.user_info.get("staffNumber") or username),
            role=str(self.user_info.get("role") or ""),
        )

    def get_bugs_by_plan(self, plan_version: str) -> list[dict[str, Any]]:
        plan_version = plan_version.strip()
        if not plan_version:
            raise BugClientError("请输入计划版本")
        if "Authorization" not in self.session.headers:
            raise AuthenticationError("请先登录")

        try:
            response = self.session.post(
                self._url("/rest/bug/getBugsByPlan"),
                params={"planVersion": plan_version},
                timeout=self.timeout,
            )
        except requests.RequestException as exc:
            raise BugClientError(f"缺陷接口请求失败：{exc}") from exc

        if response.status_code == 401:
            raise AuthenticationError("登录状态已失效，请重新登录")
        try:
            payload = response.json()
        except ValueError as exc:
            raise BugClientError("缺陷接口返回了无法识别的数据") from exc

        if response.status_code != 200 or not payload.get("success"):
            message = payload.get("message") or payload.get("msg") or "查询缺陷失败"
            raise BugClientError(str(message))

        data = payload.get("data")
        if data is None:
            return []
        if not isinstance(data, list):
            raise BugClientError("缺陷接口的数据格式异常")
        return [item for item in data if isinstance(item, dict)]

    def get_plans_by_year(self, year: int) -> list[PlanVersion]:
        if "Authorization" not in self.session.headers:
            raise AuthenticationError("请先登录")
        try:
            response = self.session.post(
                self._url("/rest/plan/getByYear"),
                params={"year": int(year)},
                timeout=self.timeout,
            )
        except requests.RequestException as exc:
            raise BugClientError(f"计划版本接口请求失败：{exc}") from exc

        if response.status_code == 401:
            raise AuthenticationError("登录状态已失效，请重新登录")
        try:
            payload = response.json()
        except ValueError as exc:
            raise BugClientError("计划版本接口返回了无法识别的数据") from exc
        if response.status_code != 200 or not payload.get("success"):
            message = payload.get("message") or payload.get("msg") or "查询计划版本失败"
            raise BugClientError(str(message))

        data = payload.get("data") or []
        if not isinstance(data, list):
            raise BugClientError("计划版本接口的数据格式异常")
        plans: list[PlanVersion] = []
        seen: set[str] = set()
        for item in data:
            if not isinstance(item, dict):
                continue
            plan_version = str(item.get("planVersion") or "").strip()
            if not plan_version or plan_version in seen:
                continue
            try:
                plan_year = int(item.get("year") or year)
                month = int(item.get("month") or 0)
            except (TypeError, ValueError):
                continue
            if not 1 <= month <= 12:
                continue
            plans.append(
                PlanVersion(
                    year=plan_year,
                    month=month,
                    plan_version=plan_version,
                )
            )
            seen.add(plan_version)
        return plans

    def close(self) -> None:
        self.session.close()
        self.user_info = {}
