"""
Configuration for the regulation project initializer.
"""
from dataclasses import dataclass, field
from pathlib import Path


@dataclass
class GitLabConfig:
    """GitLab connection settings."""
    base_url: str = "https://gitlab.sbs-software.com"
    api_url: str = "https://gitlab.sbs-software.com/api/v4"
    template_project_path: str = "sbs-cloud/dbpp/pure/pure-parameterization/pure-param-templates"
    private_token: str = ""  # Set via env var GITLAB_TOKEN or CLI arg



@dataclass
class AppConfig:
    """Top-level application configuration."""
    gitlab: GitLabConfig = field(default_factory=GitLabConfig)
    output_base_dir: Path = Path("./generated_projects")
    project_base_name: str = ""
    project_type: str = "both"  # both | regulation | output
    regulation_folder_name: str = "regulation"
    output_folder_name: str = "output"
    regulation_file_map: dict[str, str] = field(default_factory=dict)
    output_file_map: dict[str, str] = field(default_factory=dict)
    dictionary_identification: dict[str, str] = field(default_factory=dict)
    regulation_identification: dict[str, str] = field(default_factory=dict)
    output_identification: dict[str, str] = field(default_factory=dict)
    transformation_rules_identification: dict[str, str] = field(default_factory=dict)
    transformed_data_identification: dict[str, str] = field(default_factory=dict)
    ui_dashboard_identification: dict[str, str] = field(default_factory=dict)
    output_structure_identification: dict[str, str] = field(default_factory=dict)
    regulation_project_name: str = ""   # extracted from Jira, e.g. "pure-regulation-frBdfBdp"
    output_project_name: str = ""       # extracted from Jira, e.g. "pure-regulatoryoutput-frBdfBdp-bdfBdpRpc"
