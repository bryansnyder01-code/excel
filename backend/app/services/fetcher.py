"""Live data: fetch a CSV, a public Google Sheet or a JSON API into columns and rows.

Security (the server fetches URLs that users type, so it must not become a way into the private network):
* only http and https;
* the host is resolved and every address must be public: private, loopback, link-local, multicast,
  reserved and unspecified addresses are refused, for the first URL and after every redirect (at most 3);
* 10 second limit for the whole fetch, 5 MB body limit, 20,000 rows and 100 columns.

Known gap: the address is checked, then httpx resolves the name again to connect, so a DNS server that
answers differently the second time (DNS rebinding) is not fully ruled out. Run the backend without access
to sensitive internal services, or put an egress proxy in front, if that matters for your deployment.

Errors are `FetchError` with a plain-English message that is safe to show users.
"""

from __future__ import annotations

import asyncio
import csv
import io
import ipaddress
import json
import re
import socket
import time
from dataclasses import dataclass, field
from urllib.parse import parse_qs, urljoin, urlparse

import httpx

from app.schemas import NumberFormat

TIMEOUT = 10.0
MAX_BYTES = 5 * 1024 * 1024
MAX_REDIRECTS = 3
MAX_ROWS = 20_000
MAX_COLS = 100
USER_AGENT = "SheetCraft-AI/1.0 (+live data)"

GSHEET_PRIVATE = ("That Google Sheet isn't public. In Google Sheets, click Share, set General access to "
                  "\"Anyone with the link\", then try again.")


class FetchError(ValueError):
    """The data couldn't be fetched or read. The message is safe to show users."""


@dataclass
class FetchResult:
    columns: list[str]
    rows: list[list[str | float | None]]
    formats: list[NumberFormat] = field(default_factory=list)
    fetched_at: int = 0  # ms since epoch
    truncated: bool = False


# ---------------------------------------------------------------------------
# URL checks
# ---------------------------------------------------------------------------


async def resolve_host(host: str, port: int) -> list[str]:
    """Every address the host resolves to. Tests monkeypatch this."""
    loop = asyncio.get_running_loop()
    infos = await loop.getaddrinfo(host, port, type=socket.SOCK_STREAM)
    return [info[4][0] for info in infos]


def _blocked(ip: str) -> bool:
    try:
        addr = ipaddress.ip_address(ip.split("%")[0])
    except ValueError:
        return True
    if isinstance(addr, ipaddress.IPv6Address) and addr.ipv4_mapped:
        addr = addr.ipv4_mapped
    return (addr.is_private or addr.is_loopback or addr.is_link_local or addr.is_multicast
            or addr.is_reserved or addr.is_unspecified or not addr.is_global)


async def check_url(url: str) -> None:
    """Raise FetchError unless `url` is http(s) and its host resolves only to public addresses."""
    try:
        parsed = urlparse(url)
        port = parsed.port or (443 if parsed.scheme == "https" else 80)
    except ValueError:
        raise FetchError("That web address isn't valid.") from None
    if parsed.scheme not in ("http", "https"):
        raise FetchError("Only http:// and https:// addresses can be used.")
    host = parsed.hostname
    if not host:
        raise FetchError("That web address has no host name.")
    if parsed.username or parsed.password:
        raise FetchError("Addresses with a user name or password in them can't be used.")
    try:
        ipaddress.ip_address(host)
        addresses = [host]
    except ValueError:
        try:
            addresses = await resolve_host(host, port)
        except (OSError, UnicodeError):
            raise FetchError(f"Couldn't find the server \"{host}\". Check the address.") from None
    if not addresses or any(_blocked(a) for a in addresses):
        raise FetchError("That address points to a private or local network, which isn't allowed.")


_GSHEET = re.compile(r"^/spreadsheets/d/([A-Za-z0-9_-]{20,})")


def gsheet_csv_url(url: str) -> str:
    """A Google Sheets share/edit link -> its CSV export link (the tab in gid=, else the first tab)."""
    url = url.strip()
    parsed = urlparse(url)
    if parsed.hostname not in ("docs.google.com",):
        raise FetchError("That isn't a Google Sheets link. It should start with https://docs.google.com/spreadsheets/.")
    if parsed.path.startswith("/spreadsheets/d/e/"):  # "Publish to the web" link
        base = parsed.path.split("/pub")[0]
        gid = parse_qs(parsed.query).get("gid", [None])[0]
        return f"https://docs.google.com{base}/pub?output=csv" + (f"&gid={gid}" if gid else "")
    m = _GSHEET.match(parsed.path)
    if not m:
        raise FetchError("That isn't a Google Sheets link. It should look like https://docs.google.com/spreadsheets/d/...")
    gid = parse_qs(parsed.query).get("gid", [None])[0] or parse_qs(parsed.fragment).get("gid", [None])[0]
    if gid is not None and not gid.isdigit():
        gid = None
    return f"https://docs.google.com/spreadsheets/d/{m.group(1)}/export?format=csv&gid={gid or 0}"


# ---------------------------------------------------------------------------
# Fetching
# ---------------------------------------------------------------------------


def make_client() -> httpx.AsyncClient:
    return httpx.AsyncClient(timeout=TIMEOUT, follow_redirects=False, headers={"User-Agent": USER_AGENT})


async def _get(url: str, client: httpx.AsyncClient) -> tuple[int, str, bytes]:
    """GET with manual redirects (each target re-checked) and a size cap. Returns (status, content type, body)."""
    for _ in range(MAX_REDIRECTS + 1):
        await check_url(url)
        async with client.stream("GET", url, follow_redirects=False) as resp:
            if resp.status_code in (301, 302, 303, 307, 308) and resp.headers.get("location"):
                url = urljoin(url, resp.headers["location"])
                continue
            if int(resp.headers.get("content-length") or 0) > MAX_BYTES:
                raise FetchError("That file is larger than 5 MB.")
            body = bytearray()
            async for chunk in resp.aiter_bytes():
                body += chunk
                if len(body) > MAX_BYTES:
                    raise FetchError("That file is larger than 5 MB.")
            return resp.status_code, resp.headers.get("content-type", ""), bytes(body)
    raise FetchError("That address redirects too many times.")


async def fetch_table(kind: str, url: str, path: str | None = None,
                      client: httpx.AsyncClient | None = None) -> FetchResult:
    """Fetch and parse. `kind` is csv, gsheet or json."""
    url = (url or "").strip()
    if not url:
        raise FetchError("Enter a web address.")
    if len(url) > 2000:
        raise FetchError("That web address is too long.")
    if kind == "gsheet":
        url = gsheet_csv_url(url)
    elif kind not in ("csv", "json"):
        raise FetchError("The source must be csv, gsheet or json.")

    own = client is None
    client = client or make_client()
    try:
        status, ctype, body = await asyncio.wait_for(_get(url, client), TIMEOUT)
    except asyncio.TimeoutError:
        raise FetchError("The server took longer than 10 seconds to answer.") from None
    except httpx.HTTPError:
        raise FetchError("Couldn't connect to that server. Check the address and try again.") from None
    finally:
        if own:
            await client.aclose()

    html = "html" in ctype.lower() or body[:200].lstrip().lower().startswith((b"<!doctype html", b"<html"))
    if kind == "gsheet" and (status in (401, 403) or (status == 200 and html)):
        raise FetchError(GSHEET_PRIVATE)
    if kind == "gsheet" and status == 404:
        raise FetchError("That Google Sheet (or tab) wasn't found. Check the link.")
    if status in (401, 403):
        raise FetchError(f"The server refused access ({status}). The data needs to be public.")
    if status == 404:
        raise FetchError("Nothing was found at that address (404).")
    if status >= 400:
        raise FetchError(f"The server answered with an error ({status}). Try again later.")
    if html:
        raise FetchError("That address returned a web page, not data. Use a direct link to the CSV or JSON.")

    result = parse_json(body, path) if kind == "json" else parse_csv(body)
    result.fetched_at = int(time.time() * 1000)
    return result


# ---------------------------------------------------------------------------
# Parsing
# ---------------------------------------------------------------------------


def _decode(body: bytes) -> str:
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            return body.decode(enc)
        except UnicodeDecodeError:
            continue
    return body.decode("utf-8", "replace")


def _column_formats(rows: list[list], ncols: int, votes: list[dict]) -> list[NumberFormat]:
    return [max(v, key=v.get) if v else NumberFormat.text for v in votes[:ncols]]


def parse_csv(body: bytes) -> FetchResult:
    from app.services.importer import _parse_csv_value

    text = _decode(body)
    if not text.strip():
        raise FetchError("That file is empty.")
    try:
        dialect = csv.Sniffer().sniff(text[:4096], delimiters=",;\t|")
    except csv.Error:
        dialect = csv.excel
    try:
        table = [r for r in csv.reader(io.StringIO(text), dialect) if any(c.strip() for c in r)]
    except csv.Error:
        raise FetchError("That file couldn't be read as CSV.") from None
    if not table:
        raise FetchError("That file is empty.")
    truncated = len(table) - 1 > MAX_ROWS or max(len(r) for r in table) > MAX_COLS
    header, body_rows = table[0], table[1 : MAX_ROWS + 1]
    ncols = min(MAX_COLS, max(len(r) for r in table[: MAX_ROWS + 1]))
    columns = [(header[i].strip() if i < len(header) else "") or f"Column {i + 1}" for i in range(ncols)]
    votes: list[dict] = [{} for _ in range(ncols)]
    rows = []
    for raw in body_rows:
        out = []
        for i in range(ncols):
            value, fmt = _parse_csv_value(raw[i]) if i < len(raw) else (None, NumberFormat.text)
            if isinstance(value, str) and value.startswith("="):
                value = "'" + value  # data, never a formula
            out.append(value)
            if value is not None:
                votes[i][fmt] = votes[i].get(fmt, 0) + 1
        rows.append(out)
    return FetchResult(columns, rows, _column_formats(rows, ncols, votes), truncated=truncated)


def _at_path(data, path: str | None):
    for part in [p for p in (path or "").split(".") if p.strip()]:
        part = part.strip()
        if isinstance(data, dict) and part in data:
            data = data[part]
        elif isinstance(data, list) and part.lstrip("-").isdigit() and -len(data) <= int(part) < len(data):
            data = data[int(part)]
        else:
            raise FetchError(f'The path "{path}" wasn\'t found in the data (stopped at "{part}").')
    return data


def _find_list(data, depth: int = 4) -> list[str] | None:
    """Dotted path (as parts) to the first list inside nested objects, breadth first."""
    queue = [([], data)]
    while queue:
        parts, node = queue.pop(0)
        if isinstance(node, list) and parts:
            return parts
        if isinstance(node, dict) and len(parts) < depth:
            queue += [(parts + [str(k)], v) for k, v in node.items()]
    return None


def _json_value(v) -> str | float | None:
    if v is None:
        return None
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        return "'" + v if v.startswith("=") else v
    return json.dumps(v, ensure_ascii=False)


def parse_json(body: bytes, path: str | None = None) -> FetchResult:
    try:
        data = json.loads(_decode(body))
    except ValueError:
        raise FetchError("That address didn't return valid JSON.") from None
    data = _at_path(data, path)
    if isinstance(data, dict):
        found = _find_list(data)
        hint = f' Try the path "{".".join(([path] if path else []) + found)}".' if found else ""
        raise FetchError("The JSON is an object, not a list of records. Set the path to the list." + hint)
    if not isinstance(data, list):
        raise FetchError("The JSON at that path isn't a list of records.")
    truncated = len(data) > MAX_ROWS
    records = data[:MAX_ROWS]
    columns: list[str] = []
    seen: set[str] = set()
    for rec in records:
        keys = rec.keys() if isinstance(rec, dict) else ["value"]
        for k in keys:
            if k not in seen:
                if len(columns) >= MAX_COLS:
                    truncated = True
                    break
                seen.add(k)
                columns.append(str(k))
    if not columns:
        raise FetchError("The list is empty.")
    votes: list[dict] = [{} for _ in columns]
    rows = []
    for rec in records:
        rec = rec if isinstance(rec, dict) else {"value": rec}
        out = []
        for i, k in enumerate(columns):
            v = _json_value(rec.get(k))
            out.append(v)
            if isinstance(v, float):
                fmt = NumberFormat.integer if v.is_integer() else NumberFormat.decimal
                votes[i][fmt] = votes[i].get(fmt, 0) + 1
            elif v is not None:
                votes[i][NumberFormat.text] = votes[i].get(NumberFormat.text, 0) + 1
        rows.append(out)
    return FetchResult(columns, rows, _column_formats(rows, len(columns), votes), truncated=truncated)
