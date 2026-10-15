"""
Parser for Jira ticket export files (HTML/DOC and PDF formats).

Extracts regulation project information from PURE_RQM Jira tickets
to auto-fill the project initialization wizard.
"""
import re
import logging
from dataclasses import dataclass, field
from pathlib import Path

logger = logging.getLogger(__name__)


@dataclass
class JiraTicketData:
    """Extracted data from a Jira ticket."""
    ticket_key: str = ""
    title: str = ""
    project_type: str = "regulation"  # "regulation", "output", "both"
    regulation_domain_code: str = ""  # e.g. "euEbaRp"
    regulation_code_lower: str = ""   # e.g. "euebarp"
    regulator: str = ""               # e.g. "eba"
    regulator_subgroup: str = ""      # e.g. "rp"
    version: str = ""                 # e.g. "202512.1.0"
    validity_date: str = ""           # e.g. "31/12/2025"
    regulator_code: str = ""          # e.g. "euebarp" (for regulation sheet)
    authority_name: str = ""          # e.g. "EBA - European Banking Authority"
    documentation_source: str = ""    # URL
    regulator_name: str = ""          # e.g. regulation name placeholder
    git_project_url: str = ""         # e.g. "https://innersource.../eba/rp"
    linked_tickets: list = field(default_factory=list)
    file_mappings: dict = field(default_factory=dict)  # template_name -> target_name
    # ── Output-specific fields ──
    output_code: str = ""              # e.g. "ebaResol1"
    output_name: str = ""              # e.g. "EBA RESOL1"
    output_description: str = ""       # e.g. "RESOL1"
    output_periodicity: str = ""       # e.g. "Yearly"
    output_type: str = ""              # e.g. "dimensional"
    common_collection: str = ""        # e.g. ""
    submission_format: str = ""        # e.g. "xbrl-csv"
    container_type: str = ""           # e.g. "directory"
    # ── Extracted project folder names ──
    regulation_project_name: str = ""   # e.g. "pure-regulation-frBdfBdp"
    output_project_name: str = ""       # e.g. "pure-regulatoryoutput-frBdfBdp-bdfBdpRpc"


def parse_jira_ticket(file_path: str) -> JiraTicketData:
    """
    Parse a Jira ticket export file (HTML with .doc/.html extension).
    Returns a JiraTicketData with extracted information.
    """
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"Fichier introuvable: {file_path}")

    # Try UTF-8 first, fall back to latin-1
    try:
        content = path.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        content = path.read_text(encoding="latin-1")

    data = JiraTicketData()

    # ── Ticket key ──
    key_match = re.search(r'\[#?(PURE_RQM-\d+)\]', content)
    if key_match:
        data.ticket_key = key_match.group(1)

    # ── Title / summary ──
    summary_match = re.search(
        r'\[(?:PURE_RQM-\d+)\]\s*(?:&nbsp;)?\s*<a[^>]*>(.*?)</a>', content
    )
    if summary_match:
        data.title = _clean_html(summary_match.group(1))

    # ── Project type from title ──
    title_lower = data.title.lower()
    has_regulation = bool(re.search(r'\bregulation\b', title_lower))
    has_output = bool(re.search(r'\boutput\b', title_lower))
    if has_regulation and has_output:
        data.project_type = "both"
    elif has_output:
        data.project_type = "output"
    else:
        data.project_type = "regulation"

    # ── Linked tickets (for context) ──
    linked_matches = re.findall(
        r'<a\s+href="[^"]*browse/(PURE_RQM-\d+)"[^>]*>\s*'
        r'(?:<strike>)?(PURE_RQM-\d+)(?:</strike>)?\s*</a>\s*</td>\s*<td>\s*(.*?)\s*</td>',
        content, re.DOTALL,
    )
    for _, key, desc in linked_matches:
        desc_clean = _clean_html(desc).strip()
        data.linked_tickets.append({"key": key, "description": desc_clean})

    # ── Extract description area ──
    desc_match = re.search(r'<td\s+id="descriptionArea">(.*?)</td>', content, re.DOTALL)
    description = desc_match.group(1) if desc_match else content

    # ── Regulation domain code from title [xxxx] pattern ──
    domain_from_title = re.search(r'\[(eu\w+)\]', data.title, re.IGNORECASE)
    if domain_from_title:
        data.regulation_domain_code = domain_from_title.group(1)

    # ── Regulation domain code from project name pattern ──
    project_name_match = re.search(
        r'pure-regulation-<font\s+color="#4c9aff">\s*<b>(.*?)</b>',
        description,
    )
    if project_name_match:
        data.regulation_domain_code = project_name_match.group(1).strip()

    # ── Regulator (eba, ecb, etc.) ──
    regulator_match = re.search(
        r'(?:corresponding to regulator|subgroup\s*:\s*")\s*'
        r'(?:\()?<font\s+color="#4c9aff">\s*<b>(.*?)</b>',
        description, re.DOTALL,
    )
    if regulator_match:
        data.regulator = regulator_match.group(1).strip()

    # ── Regulator subgroup ──
    subgroup_match = re.search(
        r'regulator subgroup.*?<font\s+color="#4c9aff">\s*<b>(.*?)</b>',
        description, re.DOTALL,
    )
    if subgroup_match:
        data.regulator_subgroup = subgroup_match.group(1).strip()

    # ── Version ──
    version_match = re.search(
        r'Set version with\s*<font\s+color="#4c9aff">\s*<b>(.*?)</b>',
        description, re.DOTALL,
    )
    if version_match:
        data.version = version_match.group(1).strip()

    # ── Validity starting date ──
    validity_match = re.search(
        r'Set validityStartingDate with\s*<font\s+color="#4c9aff">\s*<b>(.*?)</b>',
        description, re.DOTALL,
    )
    if validity_match:
        data.validity_date = validity_match.group(1).strip()

    # ── Regulator code (for identification sheet) ──
    reg_code_match = re.search(
        r'Set regulatorCode with\s*<font\s+color="#4c9aff">\s*<b>(.*?)</b>',
        description, re.DOTALL,
    )
    if reg_code_match:
        data.regulator_code = reg_code_match.group(1).strip()
    else:
        data.regulator_code = data.regulation_domain_code.lower()

    data.regulation_code_lower = data.regulation_domain_code.lower()

    # ── Authority / author ──
    author_match = re.search(
        r'Set author with\s*"([^"]+)"',
        description,
    )
    if author_match:
        data.authority_name = author_match.group(1).strip()

    # ── Documentation source URL ──
    doc_source_match = re.search(
        r'Set documentationSource with\s*<a[^>]+href="([^"]+)"',
        description,
    )
    if doc_source_match:
        data.documentation_source = doc_source_match.group(1).strip()
    else:
        # Fallback: look for documentationSource context
        doc_sect = re.search(
            r'documentationSource.*?<a[^>]+href="([^"]+)"',
            description, re.DOTALL,
        )
        if doc_sect:
            data.documentation_source = doc_sect.group(1).strip()

    # ── Regulator name / description placeholder ──
    name_match = re.search(
        r'Fill name and description with\s*<font\s+color="#4c9aff">\s*<b>(.*?)</b>',
        description, re.DOTALL,
    )
    if name_match:
        raw_name = name_match.group(1).strip()
        # Clean trailing {} artifacts
        raw_name = re.sub(r'[{}]+$', '', raw_name).strip()
        data.regulator_name = raw_name

    # ── Git project URL ──
    url_match = re.search(
        r'has to be created in the folder\s*:\s*<a[^>]+href="([^"]+)"',
        description, re.DOTALL,
    )
    if url_match:
        data.git_project_url = url_match.group(1).strip()

    # ── Extract file rename mappings ──
    # Pattern: pure-TYPE-<b>templateCode</b>.EXT to pure-TYPE-<font...><b>targetCode</b></font>.EXT
    file_mapping_matches = re.findall(
        r'(pure-[\w]+-)<b>([^<]+)</b>\.(xlsx|xlsm)\s+to\s+'
        r'(pure-[\w]+-)<font[^>]*>\s*<b>([^<]+)</b>\s*</font>\.(xlsx|xlsm)',
        description, re.IGNORECASE,
    )
    for prefix_src, placeholder, ext_src, prefix_dst, target_code, ext_dst in file_mapping_matches:
        src_name = f"{prefix_src}{placeholder}.{ext_src}"
        dst_name = f"{prefix_dst}{target_code}.{ext_dst}"
        data.file_mappings[src_name] = dst_name

    # ── Also detect project type from description content ──
    has_dict = bool(re.search(r'pure-dictionary-', description))
    has_reg_file = bool(re.search(r'pure-regulation-', description))
    has_output_file = bool(re.search(r'pure-regulatoryOutput-', description))
    has_output_struct = bool(re.search(r'pure-outputStructure-', description))

    if (has_dict or has_reg_file) and (has_output_file or has_output_struct):
        data.project_type = "both"
    elif has_output_file or has_output_struct:
        data.project_type = "output"
    elif has_dict or has_reg_file:
        data.project_type = "regulation"

    logger.info(
        "Parsed Jira ticket: %s → code=%s, version=%s, type=%s",
        data.ticket_key, data.regulation_domain_code, data.version, data.project_type,
    )
    return data


def _clean_html(text: str) -> str:
    """Remove HTML tags from text."""
    return re.sub(r'<[^>]+>', '', text).strip()


# ══════════════════════════════════════════════════════════
#  PDF Parser — extracts data from Jira/Xray PDF exports
# ══════════════════════════════════════════════════════════

def parse_jira_pdf(file_path: str) -> JiraTicketData:
    """
    Parse a Jira/Xray PDF export file.
    Extracts all regulation project information from the plain text.
    Returns a JiraTicketData with extracted information.
    """
    try:
        import pymupdf
    except ImportError:
        raise ImportError(
            "Le module 'pymupdf' est requis pour lire les PDF. "
            "Installer avec: pip install pymupdf"
        )

    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"Fichier introuvable: {file_path}")

    # Extract all text from PDF, normalize whitespace for reliable regex
    doc = pymupdf.open(str(path))
    raw_pages = []
    for page in doc:
        raw_pages.append(page.get_text())
    doc.close()
    full_text_raw = "\n".join(raw_pages)
    # Replace bullet characters (uf0a7, etc.) with spaces
    full_text_raw = re.sub(r'[\uf0a7\uf0b7\u2022\u2023\u25e6]', ' ', full_text_raw)
    # Join lines that are part of the same sentence (line ending without period/colon)
    # but keep paragraph breaks (double newlines)
    full_text = re.sub(r'(?<!\n)\n(?!\n)', ' ', full_text_raw)
    # Also rejoin hyphenated line breaks (e.g. "risk-and-\ndata" -> "risk-and-data")
    full_text = re.sub(r'-\s+', '-', full_text)
    # Collapse multiple spaces
    full_text = re.sub(r' {2,}', ' ', full_text)

    data = JiraTicketData()

    # ── Ticket key (e.g. PURE_RQM-5363) ──
    key_match = re.search(r'(PURE_RQM-\d+)', full_text)
    if key_match:
        data.ticket_key = key_match.group(1)

    # ── Title from "[xxx] - GIT ... project creation/initialization" pattern ──
    title_match = re.search(
        r'\[(\w+)\]\s*[-–—]\s*GIT\s+(.*?)(?:project\s+(?:creation|initialization)|$)',
        full_text, re.IGNORECASE,
    )
    if title_match:
        data.regulation_domain_code = title_match.group(1)
        data.title = f"[{title_match.group(1)}] - GIT {title_match.group(2).strip()} project"
    else:
        # Fallback: look for [xxxYyy] pattern anywhere
        domain_match = re.search(r'\[(\w+)\]', full_text, re.IGNORECASE)
        if domain_match:
            data.regulation_domain_code = domain_match.group(1)

    data.regulation_code_lower = data.regulation_domain_code.lower()

    # ── Project type from title ──
    title_lower = data.title.lower()
    has_regulation = bool(re.search(r'\bregulation\b', title_lower))
    has_output = bool(re.search(r'\boutput\b', title_lower))
    if has_regulation and has_output:
        data.project_type = "both"
    elif has_output:
        data.project_type = "output"
    else:
        data.project_type = "regulation"

    # ── Git project URL (folder where project must be created) ──
    url_match = re.search(
        r'has to be created in the folder\s*:\s*\n?\s*(https?://\S+)',
        full_text,
    )
    if url_match:
        data.git_project_url = url_match.group(1).strip()

    # ── Regulator from "corresponding to regulator (xxx)" ──
    regulator_match = re.search(
        r'corresponding to regulator\s*\((\w+)\)', full_text, re.IGNORECASE,
    )
    if regulator_match:
        data.regulator = regulator_match.group(1).strip().lower()

    # ── Regulator subgroup from "regulator subgroup (xxx)" ──
    subgroup_match = re.search(
        r'regulator subgroup\s*\((\w+)\)', full_text, re.IGNORECASE,
    )
    if subgroup_match:
        data.regulator_subgroup = subgroup_match.group(1).strip().lower()

    # ── File rename mappings ──
    # Pattern: "pure-TYPE-templateCode.xlsx to pure-TYPE-targetCode.xlsx"
    file_mapping_matches = re.findall(
        r'(pure-[\w]+-)([\w.()\-]+)\.(xlsx|xlsm)\s+to\s+(pure-[\w]+-)([\w.()\-]+)\.(xlsx|xlsm)',
        full_text, re.IGNORECASE,
    )
    for prefix_src, placeholder, ext_src, prefix_dst, target_code, ext_dst in file_mapping_matches:
        src_name = f"{prefix_src}{placeholder}.{ext_src}"
        dst_name = f"{prefix_dst}{target_code}.{ext_dst}"
        data.file_mappings[src_name] = dst_name

    # ── Version from "Set version with XXX" ──
    version_match = re.search(
        r'Set version with\s+(\S+)', full_text,
    )
    if version_match:
        data.version = version_match.group(1).strip()

    # ── Validity date from "Set validityStartingDate with XXX" ──
    validity_match = re.search(
        r'Set validityStartingDate with\s+(\S+)', full_text,
    )
    if validity_match:
        data.validity_date = validity_match.group(1).strip()

    # ── Regulator code from "Set regulatorCode with XXX" ──
    reg_code_match = re.search(
        r'Set regulatorCode with\s+(\S+)', full_text,
    )
    if reg_code_match:
        data.regulator_code = reg_code_match.group(1).strip()
    else:
        data.regulator_code = data.regulation_domain_code.lower()

    # ── Author from 'Set author with "XXX"' (handles smart quotes too) ──
    author_match = re.search(
        r'Set author with\s+["\u201c]([^"\u201d]+)["\u201d]', full_text,
    )
    if author_match:
        data.authority_name = author_match.group(1).strip()

    # ── Documentation source URL (may span multiple lines in PDF) ──
    doc_source_match = re.search(
        r'Set documentationSource with\s+(https?://[^\s]+(?:[-/][^\s]+)*)',
        full_text,
    )
    if doc_source_match:
        raw_url = doc_source_match.group(1).strip()
        # Remove trailing non-URL chars (bullets, letters stuck to end)
        raw_url = re.sub(r'[^a-zA-Z0-9/_\-.:?&=%#]+$', '', raw_url)
        data.documentation_source = raw_url

    # ── Regulator name from "Fill name and description with XXX" ──
    name_match = re.search(
        r'Fill name and description with\s+(\S+)', full_text,
    )
    if name_match:
        raw_name = name_match.group(1).strip()
        # Clean trailing {} artifacts from PDF extraction
        raw_name = re.sub(r'[{}]+$', '', raw_name).strip()
        data.regulator_name = raw_name

    # ══════════════════════════════════════════
    # Output-specific fields (different format)
    # ══════════════════════════════════════════

    # ── Version from "Version = XXX" (output format) ──
    if not data.version:
        ver2 = re.search(r'Version\s*=\s*(\S+)', full_text)
        if ver2:
            data.version = ver2.group(1).strip()

    # ── Validity from "Validity starting date = XXX" ──
    if not data.validity_date:
        val2 = re.search(r'Validity\s+starting\s+date\s*=\s*(\S+)', full_text, re.IGNORECASE)
        if val2:
            data.validity_date = val2.group(1).strip()

    # ── Output code from "Code = XXX" ──
    code_match = re.search(r'(?:^|\s)Code\s*=\s*(\S+)', full_text)
    if code_match:
        data.output_code = code_match.group(1).strip()

    # ── Output name from "Name = XXX" (may be multi-word) ──
    name_out = re.search(r'(?:^|\s)Name\s*=\s*(.+?)(?:\s*Description|$)', full_text)
    if name_out:
        data.output_name = name_out.group(1).strip()

    # ── Output description from "Description = XXX" ──
    desc_out = re.search(r'(?:^|\s)Description\s*=\s*(\S+)', full_text)
    if desc_out:
        data.output_description = desc_out.group(1).strip()

    # ── Periodicity from "Periodicity = XXX" ──
    period = re.search(r'Periodicity\s*=\s*(\S+)', full_text, re.IGNORECASE)
    if period:
        data.output_periodicity = period.group(1).strip()

    # ── Output type from "Set the outputType with XXX" ──
    otype = re.search(r'Set the outputType with\s+(\S+)', full_text, re.IGNORECASE)
    if otype:
        data.output_type = otype.group(1).strip()

    # ── Common collection from "Set the commonCollection with XXX" ──
    # Value may be empty in the ticket — only capture if followed by a real value
    ccol = re.search(r'Set the commonCollection with\s+(?!Set\b|Sheet\b)(\S+)', full_text, re.IGNORECASE)
    if ccol:
        data.common_collection = ccol.group(1).strip()

    # ── Submission format from "Set the submissionFormat with XXX" ──
    sfmt = re.search(r'Set the submissionFormat with\s+(\S+)', full_text, re.IGNORECASE)
    if sfmt:
        data.submission_format = sfmt.group(1).strip()

    # ── Container type from "Set the containerType with XXX" ──
    ctype = re.search(r'Set the containerType with\s+(\S+)', full_text, re.IGNORECASE)
    if ctype:
        data.container_type = ctype.group(1).strip()

    # ── Extract project folder names from description ──
    # Try "Regulation : pure-regulation-XXX" first, then standalone bullet "pure-regulation-XXX"
    reg_proj = re.search(r'Regulation\s*:\s*(pure-regulation-\S+)', full_text, re.IGNORECASE)
    if reg_proj:
        data.regulation_project_name = reg_proj.group(1).strip()
    else:
        # Match standalone project name (bullet point), exclude file references (.xlsx/.xlsm)
        reg_proj_alt = re.findall(r'(pure-regulation-[a-zA-Z0-9]+)(?!\S*\.xls)', full_text)
        if reg_proj_alt:
            data.regulation_project_name = reg_proj_alt[0].strip()

    out_proj = re.search(r'Regulatory\s+output\s*:\s*(pure-regulatoryoutput-\S+)', full_text, re.IGNORECASE)
    if out_proj:
        data.output_project_name = out_proj.group(1).strip()
    else:
        # Match standalone project name, exclude file references (.xlsx/.xlsm)
        out_proj_alt = re.findall(r'(pure-regulatoryoutput-[a-zA-Z0-9-]+)(?!\S*\.xls)', full_text)
        if out_proj_alt:
            data.output_project_name = out_proj_alt[0].strip()

    # ── Detect project type from file mentions ──
    has_dict = bool(re.search(r'pure-dictionary-', full_text))
    has_reg_file = bool(re.search(r'pure-regulation-', full_text))
    has_output_file = bool(re.search(r'pure-regulatoryOutput-', full_text))
    has_output_struct = bool(re.search(r'pure-outputStructure-', full_text))

    if (has_dict or has_reg_file) and (has_output_file or has_output_struct):
        data.project_type = "both"
    elif has_output_file or has_output_struct:
        data.project_type = "output"
    elif has_dict or has_reg_file:
        data.project_type = "regulation"

    logger.info(
        "Parsed Jira PDF: %s → code=%s, version=%s, type=%s",
        data.ticket_key, data.regulation_domain_code, data.version, data.project_type,
    )
    return data


def parse_jira_file(file_path: str) -> JiraTicketData:
    """
    Auto-detect file format (PDF or HTML/DOC) and parse accordingly.
    """
    path = Path(file_path)
    ext = path.suffix.lower()
    if ext == ".pdf":
        return parse_jira_pdf(file_path)
    return parse_jira_ticket(file_path)


def jira_data_to_collected(data: JiraTicketData) -> dict:
    """
    Convert parsed Jira ticket data into the 'collected' dict format
    used by the GUI and project initializer.
    """
    code = data.regulation_domain_code.upper()
    code_lower = data.regulation_code_lower or data.regulation_domain_code.lower()

    result = {
        "regulation_code": code,
        "project_type": data.project_type,
        "git_project_name": code_lower,
        "jira_ticket_key": data.ticket_key,
        "jira_title": data.title,
        "regulation_project_name": data.regulation_project_name,
        "output_project_name": data.output_project_name,
        "git_project_url": data.git_project_url,
    }

    # ── Pre-fill regulation identification ──
    if data.version or data.validity_date or data.authority_name:
        result["regulation_identification"] = {
            "regulatorCode": data.regulator_code or code_lower,
            "version": data.version,
            "validityStartingDate": data.validity_date,
            "name (FR)": data.regulator_name or "",
            "name (EN)": data.regulator_name or "",
            "description (EN)": data.regulator_name or "",
            "description (FR)": data.regulator_name or "",
            "author": data.authority_name,
            "documentationSource": data.documentation_source,
        }

    # ── Pre-fill dictionary identification ──
    if data.version or data.validity_date:
        result["dictionary_identification"] = {
            "version": data.version,
            "validityStartingDate": data.validity_date,
            "description (EN)": data.regulator_name or "",
            "description (FR)": data.regulator_name or "",
            "regulatorName (EN)": data.regulator_name or "",
            "regulatorName (FR)": data.regulator_name or "",
        }

    # ── Pre-fill other identification sheets ──
    for key in [
        "transformation_rules_identification",
        "transformed_data_identification",
        "ui_dashboard_identification",
    ]:
        if data.version or data.validity_date:
            result[key] = {
                "version": data.version,
                "validityStartingDate": data.validity_date,
                "description (EN)": "",
                "description (FR)": "",
            }

    # ── Pre-fill output identification (if output project) ──
    if data.project_type in ("both", "output"):
        if data.version or data.validity_date or data.output_code:
            result["output_identification"] = {
                "code": data.output_code,
                "version": data.version,
                "validityStartingDate": data.validity_date,
                "periodicity": data.output_periodicity,
                "commonCollection": data.common_collection,
                "outputType": data.output_type,
                "name (EN)": data.output_name,
                "name (FR)": data.output_name,
                "description (EN)": data.output_description,
                "description (FR)": data.output_description,
                "submissionFormat": data.submission_format,
                "containerType": data.container_type,
            }
            result["output_structure_identification"] = {
                "version": data.version,
                "validityStartingDate": data.validity_date,
                "name (EN)": data.output_name,
                "name (FR)": data.output_name,
                "description (EN)": data.output_description,
                "description (FR)": data.output_description,
            }

    # ── Pre-fill file mappings (explicit from ticket + auto-derived) ──
    result["file_mappings"] = data.file_mappings

    return result


def format_jira_summary(data: JiraTicketData) -> str:
    """Format a human-readable summary of the parsed Jira ticket."""
    lines = [
        "═══════════════════════════════════════════",
        "  INFORMATIONS EXTRAITES DU TICKET JIRA",
        "═══════════════════════════════════════════",
        "",
        f"  Ticket        :  {data.ticket_key}",
        f"  Titre         :  {data.title}",
        f"  Type projet   :  {data.project_type}",
        f"  Code domaine  :  {data.regulation_domain_code}",
        f"  Régulateur    :  {data.regulator}",
    ]
    if data.regulator_subgroup:
        lines.append(f"  Sous-groupe   :  {data.regulator_subgroup}")
    lines.extend([
        f"  Version       :  {data.version}",
        f"  Date validité :  {data.validity_date}",
        f"  Code régulat. :  {data.regulator_code}",
        f"  Autorité      :  {data.authority_name}",
    ])
    if data.documentation_source:
        lines.append(f"  Doc. source   :  {data.documentation_source}")
    if data.git_project_url:
        lines.append(f"  URL projet    :  {data.git_project_url}")
    if data.regulation_project_name:
        lines.append(f"  Projet reg.   :  {data.regulation_project_name}")
    if data.output_project_name:
        lines.append(f"  Projet output :  {data.output_project_name}")
    # Output-specific fields
    if data.output_code:
        lines.append(f"  Output code   :  {data.output_code}")
    if data.output_name:
        lines.append(f"  Output nom    :  {data.output_name}")
    if data.output_description:
        lines.append(f"  Output desc.  :  {data.output_description}")
    if data.output_periodicity:
        lines.append(f"  Périodicité   :  {data.output_periodicity}")
    if data.output_type:
        lines.append(f"  Output type   :  {data.output_type}")
    if data.submission_format:
        lines.append(f"  Submission fmt:  {data.submission_format}")
    if data.container_type:
        lines.append(f"  Container type:  {data.container_type}")
    if data.file_mappings:
        lines.append("")
        lines.append("  ── Fichiers à générer ──")
        for src, dst in data.file_mappings.items():
            lines.append(f"    • {src}  →  {dst}")
    if data.linked_tickets:
        lines.append("")
        lines.append("  ── Tickets liés ──")
        for lt in data.linked_tickets:
            lines.append(f"    • {lt['key']} – {lt['description']}")

    lines.append("")
    lines.append("═══════════════════════════════════════════")
    return "\n".join(lines)
