from __future__ import annotations

import re


CVE_RE = re.compile(r"(CVE-\d{4}-\d{4,7})", re.IGNORECASE)
CWE_RE = re.compile(r"(CWE-\d+)", re.IGNORECASE)
OWASP_TOP10_RE = re.compile(r"(A\d{1,2}:\d{4})", re.IGNORECASE)


def normalize_url(url: str) -> str:
    value = (url or "").strip()
    if not value:
        return ""
    if value.startswith(("http://", "https://")):
        return value
    return "https://" + value


def extract_cve(text: str) -> str | None:
    match = CVE_RE.search(text or "")
    if not match:
        return None
    return match.group(1).upper()


def extract_cwe(text: str) -> str | None:
    match = CWE_RE.search(text or "")
    if not match:
        return None
    return match.group(1).upper()


def extract_owasp_top10(text: str) -> str | None:
    match = OWASP_TOP10_RE.search(text or "")
    if not match:
        return None
    return match.group(1).upper()


def build_nvd_url(cve: str) -> str:
    return f"https://nvd.nist.gov/vuln/detail/{cve}"


def build_cwe_url(cwe: str) -> str:
    cwe_number = cwe.upper().replace("CWE-", "")
    return f"https://cwe.mitre.org/data/definitions/{cwe_number}.html"


def build_owasp_url(top10_code: str | None = None) -> str:
    if top10_code:
        return f"https://owasp.org/Top10/{top10_code}/"
    return "https://owasp.org/Top10/"


def smart_fill_reference_fields(
    reference_type: str,
    title: str,
    url: str,
) -> tuple[str, str, str]:
    ref_type = (reference_type or "Custom").strip()
    current_title = (title or "").strip()
    current_url = (url or "").strip()

    cve = extract_cve(current_title)
    cwe = extract_cwe(current_title)
    owasp_code = extract_owasp_top10(current_title)

    if ref_type in {"CVE", "NVD"} and cve:
        return cve, "NVD", build_nvd_url(cve)

    if ref_type == "CWE" and cwe:
        return cwe, "CWE", build_cwe_url(cwe)

    if ref_type == "OWASP":
        if owasp_code:
            return owasp_code, "OWASP", build_owasp_url(owasp_code)
        if current_title:
            return current_title, "OWASP", build_owasp_url()
        return "OWASP Top 10", "OWASP", build_owasp_url()

    if ref_type == "Vendor Advisory" and current_url:
        return current_title or "Vendor Advisory", ref_type, normalize_url(current_url)

    if ref_type == "Blog" and current_url:
        return current_title or "Blog Reference", ref_type, normalize_url(current_url)

    if cve:
        return cve, "NVD", build_nvd_url(cve)

    if cwe:
        return cwe, "CWE", build_cwe_url(cwe)

    if current_url:
        return current_title or "Custom Reference", ref_type, normalize_url(current_url)

    return current_title, ref_type, current_url