"""
Data models for the regulation project initializer.
"""
from dataclasses import dataclass


@dataclass
class RegulationInfo:
    """Core regulation information used to populate templates."""
    code: str                          # e.g. "COREP", "FINREP", "AnaCredit"
    version: str                       # e.g. "3.4", "2025.1"
    name: str                          # Full name, e.g. "Common Reporting Framework"
    authority: str                     # e.g. "EBA", "ECB", "BNB"
    frequency: str                     # e.g. "Quarterly", "Monthly"
    first_reporting_date: str          # e.g. "2025-03-31"
    description: str = ""
    taxonomy_version: str = ""
    taxonomy_entry_point: str = ""


@dataclass
class ProjectInput:
    """Complete input data to initialize projects."""
    regulation: RegulationInfo
