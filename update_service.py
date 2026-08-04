from __future__ import annotations

import json
import re
from dataclasses import dataclass
from typing import Optional, Tuple
from urllib.error import HTTPError, URLError
from urllib.parse import quote, unquote, urlparse
from urllib.request import Request, urlopen


APP_VERSION = "0.2.2"
GITHUB_OWNER = "a2075393995-hub"
GITHUB_REPOSITORY = "qunzhong-toupiao"
REPOSITORY_URL = f"https://github.com/{GITHUB_OWNER}/{GITHUB_REPOSITORY}"
LATEST_RELEASE_API_URL = (
    f"https://api.github.com/repos/{quote(GITHUB_OWNER, safe='')}/"
    f"{quote(GITHUB_REPOSITORY, safe='')}/releases/latest"
)
LATEST_RELEASE_PAGE_URL = f"{REPOSITORY_URL}/releases/latest"


@dataclass(frozen=True)
class ReleaseInfo:
    tag_name: str
    name: str
    html_url: str
    body: str
    published_at: str

    @property
    def version(self) -> str:
        return self.tag_name.strip().lstrip("vV")


def version_key(version: str) -> Tuple[int, int, int, int, str]:
    text = version.strip().lstrip("vV")
    match = re.match(r"^(\d+)(?:\.(\d+))?(?:\.(\d+))?(?:[-+]?(.+))?$", text)
    if not match:
        return 0, 0, 0, -1, text.lower()
    major = int(match.group(1) or 0)
    minor = int(match.group(2) or 0)
    patch = int(match.group(3) or 0)
    suffix = (match.group(4) or "").lower()
    stable_rank = 1 if not suffix else 0
    return major, minor, patch, stable_rank, suffix


def is_newer_version(candidate: str, current: str) -> bool:
    return version_key(candidate) > version_key(current)


def fetch_latest_release_from_page(timeout: float = 10.0) -> Optional[ReleaseInfo]:
    """Use GitHub's latest-release redirect when the public API is rate-limited."""
    request = Request(
        LATEST_RELEASE_PAGE_URL,
        headers={"User-Agent": f"QunzhongVote/{APP_VERSION}"},
    )
    try:
        with urlopen(request, timeout=timeout) as response:
            final_url = str(response.geturl() or "").strip()
    except HTTPError as exc:
        if exc.code == 404:
            return None
        raise RuntimeError(f"GitHub Release 页面返回 HTTP {exc.code}") from exc
    except (URLError, TimeoutError) as exc:
        reason = getattr(exc, "reason", exc)
        raise RuntimeError(f"GitHub Release 页面连接失败：{reason}") from exc

    path_parts = [unquote(part) for part in urlparse(final_url).path.split("/") if part]
    try:
        tag_index = path_parts.index("tag")
        tag_name = path_parts[tag_index + 1]
    except (ValueError, IndexError) as exc:
        raise RuntimeError("GitHub 最新版本页面没有返回有效的 Release 标签") from exc
    return ReleaseInfo(
        tag_name=tag_name,
        name=tag_name,
        html_url=final_url,
        body="",
        published_at="",
    )


def fetch_latest_release(timeout: float = 10.0) -> Optional[ReleaseInfo]:
    request = Request(
        LATEST_RELEASE_API_URL,
        headers={
            "Accept": "application/vnd.github+json",
            "User-Agent": f"QunzhongVote/{APP_VERSION}",
            "X-GitHub-Api-Version": "2022-11-28",
        },
    )
    try:
        with urlopen(request, timeout=timeout) as response:
            payload = json.load(response)
    except HTTPError as exc:
        if exc.code == 404:
            return None
        if exc.code in {403, 429}:
            return fetch_latest_release_from_page(timeout)
        raise RuntimeError(f"GitHub API 返回 HTTP {exc.code}") from exc
    except URLError as exc:
        try:
            return fetch_latest_release_from_page(timeout)
        except RuntimeError:
            reason = getattr(exc, "reason", exc)
            raise RuntimeError(f"网络连接失败：{reason}") from exc
    except TimeoutError as exc:
        raise RuntimeError("连接 GitHub 超时") from exc

    tag_name = str(payload.get("tag_name") or "").strip()
    html_url = str(payload.get("html_url") or "").strip()
    if not tag_name or not html_url:
        raise RuntimeError("GitHub Release 数据缺少版本号或下载地址")

    return ReleaseInfo(
        tag_name=tag_name,
        name=str(payload.get("name") or tag_name),
        html_url=html_url,
        body=str(payload.get("body") or ""),
        published_at=str(payload.get("published_at") or ""),
    )
