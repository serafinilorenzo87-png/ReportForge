from __future__ import annotations


FINDING_TEMPLATES = {
    "Weak SSH Credentials": {
        "severity": "High",
        "description": (
            "The target exposed an SSH service that accepted weak or guessable "
            "credentials. This issue can allow an attacker to gain unauthorized "
            "remote access to the system and move further into the environment."
        ),
        "evidence": (
            "Valid credentials were identified against the SSH service during the "
            "assessment. Authentication succeeded and interactive access was obtained."
        ),
        "remediation": (
            "Enforce strong password policies, disable weak credentials, and consider "
            "using key-based authentication with MFA where possible."
        ),
    },
    "Anonymous FTP Enabled": {
        "severity": "Medium",
        "description": (
            "The FTP service allows anonymous access. This can expose files, leak "
            "internal information, or provide an attacker with useful reconnaissance data."
        ),
        "evidence": (
            "The FTP service accepted anonymous authentication and allowed access "
            "without valid user credentials."
        ),
        "remediation": (
            "Disable anonymous FTP access unless explicitly required, and restrict "
            "access to authenticated users only."
        ),
    },
    "SMB Signing Disabled": {
        "severity": "Medium",
        "description": (
            "SMB signing was not enforced on the target. This can increase the risk "
            "of relay-style attacks in certain environments."
        ),
        "evidence": (
            "Enumeration results showed that SMB signing was supported but not required."
        ),
        "remediation": (
            "Require SMB signing on affected hosts through system configuration or "
            "domain policy where appropriate."
        ),
    },
    "Exposed Admin Interface": {
        "severity": "Medium",
        "description": (
            "An administrative or management interface was exposed to a broader network "
            "scope than necessary. This increases attack surface and can facilitate "
            "brute force, credential attacks, or exploitation of known vulnerabilities."
        ),
        "evidence": (
            "A management or administrative interface was reachable during the assessment."
        ),
        "remediation": (
            "Restrict access to trusted IP ranges, place the interface behind VPN or "
            "access controls, and reduce external exposure."
        ),
    },
    "Default Credentials": {
        "severity": "Critical",
        "description": (
            "The target accepted default or vendor-supplied credentials. This can allow "
            "immediate unauthorized access with minimal attacker effort."
        ),
        "evidence": (
            "Authentication succeeded using default or widely known credential pairs."
        ),
        "remediation": (
            "Change all default credentials immediately, review other systems for similar "
            "exposures, and implement a hardened provisioning process."
        ),
    },
    "Web Directory Listing Enabled": {
        "severity": "Low",
        "description": (
            "Directory listing was enabled on a web-accessible path. This may reveal "
            "sensitive files, internal naming conventions, or application structure."
        ),
        "evidence": (
            "A directory path returned a listing of files and folders instead of an access denial."
        ),
        "remediation": (
            "Disable directory listing on the web server and ensure sensitive files are not "
            "stored in web-accessible locations."
        ),
    },
    "Outdated Service Version": {
        "severity": "Low",
        "description": (
            "The target exposed a service version that appears outdated. Old versions may "
            "contain known vulnerabilities or unsupported components."
        ),
        "evidence": (
            "Service/version detection identified an outdated or legacy version running on the host."
        ),
        "remediation": (
            "Upgrade the affected service to a currently supported version and apply the "
            "latest security patches."
        ),
    },
    "Local File Inclusion": {
        "severity": "High",
        "description": (
            "The application appears vulnerable to Local File Inclusion (LFI). This can allow "
            "an attacker to read sensitive local files and, in some cases, escalate to code execution."
        ),
        "evidence": (
            "A crafted parameter allowed traversal or file inclusion behavior consistent with LFI."
        ),
        "remediation": (
            "Validate and sanitize file-related input, avoid dynamic file inclusion patterns, "
            "and implement strict allowlists."
        ),
    },
    "Command Injection": {
        "severity": "Critical",
        "description": (
            "The application appears vulnerable to command injection. This can allow an attacker "
            "to execute arbitrary operating system commands in the context of the application."
        ),
        "evidence": (
            "User-controlled input influenced command execution and produced behavior consistent "
            "with injected system commands."
        ),
        "remediation": (
            "Avoid shell command execution with unsanitized input, use safe APIs, validate input "
            "strictly, and apply least-privilege execution."
        ),
    },
    "SQL Injection": {
        "severity": "Critical",
        "description": (
            "The application appears vulnerable to SQL injection. This can allow unauthorized "
            "data access, modification, and possible full compromise of the underlying database."
        ),
        "evidence": (
            "Application responses were consistent with injected SQL syntax affecting query behavior."
        ),
        "remediation": (
            "Use parameterized queries, ORM protections where applicable, strict server-side validation, "
            "and least-privilege database accounts."
        ),
    },
}


def get_template_names() -> list[str]:
    return sorted(FINDING_TEMPLATES.keys())


def get_template_data(template_name: str) -> dict[str, str] | None:
    return FINDING_TEMPLATES.get(template_name)