"""
Tkinter GUI for the Regulation Project Initializer.

Replaces the console-based interactive prompts with a wizard-style interface.
"""
import logging
import os
import re
import sys
import threading
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
from pathlib import Path

# Ensure src/ is importable
sys.path.insert(0, str(Path(__file__).resolve().parent))

from config import AppConfig, GitLabConfig
from models import ProjectInput, RegulationInfo
from project_initializer import ProjectInitializer
from gitlab_publisher import GitLabPublisher
from jira_parser import parse_jira_file, jira_data_to_collected, format_jira_summary

logger = logging.getLogger("regulation-init")

# ──────────────────────────────────────────────────────────
# Helpers (duplicated from main.py to keep GUI self-contained)
# ──────────────────────────────────────────────────────────

def _sanitize_project_name(value: str) -> str:
    cleaned = re.sub(r"[^a-zA-Z0-9_-]+", "-", (value or "").strip().lower())
    return cleaned.strip("-_")


def _sanitize_branch_name(value: str) -> str:
    cleaned = re.sub(r"[^a-zA-Z0-9._/-]+", "", (value or "").strip())
    return cleaned.strip("/")


def _sanitize_file_name(value: str, extension: str) -> str:
    base = re.sub(r"[^a-zA-Z0-9._()\-]+", "-", (value or "").strip())
    base = base.strip(" .")
    if not base:
        return ""
    if not base.lower().endswith(extension.lower()):
        base = f"{base}{extension}"
    return base


def _normalize_folder_name(value: str) -> str:
    raw = (value or "").strip().lower()
    if raw in {"", "none", "aucun", ".", "./", "-"}:
        return ""
    return _sanitize_project_name(value)


def _tokenized_default_name(file_name: str, regulation_code: str) -> str:
    lower_code = _sanitize_project_name(regulation_code)
    tokenized = file_name
    replacements = {
        "regulatorCode": "REGULATION_LOWER_CODE",
        "regulationCode": "REGULATION_LOWER_CODE",
        "regulatoryDomainCode": "REGULATION_LOWER_CODE",
        "regulatoryOutputCode": "REGULATORY_OUTPUT_CODE",
        "outputStructureCode": "OUTPUT_STRUCTURE_CODE",
        "formCode": "FORM_CODE",
        "(shortName)": "SHORT_NAME",
    }
    for old, new in replacements.items():
        tokenized = tokenized.replace(old, new)
    tokenized = tokenized.replace("REGULATION_LOWER_CODE", lower_code or "regulation")
    return tokenized


def _translate_en_to_fr(text: str) -> str:
    glossary = {
        "dictionary": "dictionnaire", "for": "pour", "the": "le",
        "liquidity": "liquidite", "regulation": "reglementation",
        "output": "output", "reporting": "reporting", "framework": "cadre",
        "common": "commun", "own": "propres", "funds": "fonds",
        "test": "test", "cases": "cas", "institutions": "institutions",
    }
    parts = re.split(r"(\W+)", text)
    translated = []
    for part in parts:
        if not part or re.fullmatch(r"\W+", part):
            translated.append(part)
            continue
        lower = part.lower()
        fr = glossary.get(lower, part)
        if part[:1].isupper():
            fr = fr[:1].upper() + fr[1:]
        translated.append(fr)
    return "".join(translated)


def _translate_fr_to_en(text: str) -> str:
    glossary = {
        "dictionnaire": "dictionary", "pour": "for", "le": "the",
        "la": "the", "les": "the", "liquidite": "liquidity",
        "reglementation": "regulation", "sortie": "output",
        "rapport": "reporting", "cadre": "framework", "commun": "common",
        "propres": "own", "fonds": "funds", "test": "test",
        "cas": "cases", "institutions": "institutions",
    }
    parts = re.split(r"(\W+)", text)
    translated = []
    for part in parts:
        if not part or re.fullmatch(r"\W+", part):
            translated.append(part)
            continue
        lower = part.lower()
        en = glossary.get(lower, part)
        if part[:1].isupper():
            en = en[:1].upper() + en[1:]
        translated.append(en)
    return "".join(translated)


def _reference_dir_for_scope(scope: str) -> Path:
    src_root = Path(__file__).resolve().parent
    if scope == "regulation":
        return src_root / "regulation"
    return src_root / "Output"


def _extract_gitlab_namespace(git_project_url: str) -> str:
    """
    Extract the GitLab group namespace path from a git project URL.

    Example:
        URL = "https://innersource.soprasteria.com/sbs-cloud/dbpp/pure/pure-parameterization/bdf/bdp"
        Returns: "sbs-cloud/dbpp/pure/pure-parameterization/bdf/bdp"
    """
    if not git_project_url:
        return ""
    url = git_project_url.strip().rstrip("/")
    # Remove protocol and host
    if "://" in url:
        url = url.split("://", 1)[1]
        # Remove host part
        if "/" in url:
            url = url.split("/", 1)[1]
        else:
            return ""
    return url.strip("/")


# ──────────────────────────────────────────────────────────
# Color / style constants
# ──────────────────────────────────────────────────────────

BG = "#1e1e2e"
FG = "#cdd6f4"
ACCENT = "#6a88b8"
ACCENT_HOVER = "#74c7ec"
ENTRY_BG = "#313244"
ENTRY_FG = "#cdd6f4"
BTN_BG = "#59729a"
BTN_FG = "#d6d6ee"
HEADER_FG = "#f5c2e7"
SUCCESS_FG = "#a6e3a1"
WARN_FG = "#f9e2af"
ERR_FG = "#f38ba8"
SECTION_BG = "#B2B2B9"
BORDER_COLOR = "#45475a"
CARD_BG = "#24243a"
SURFACE_BG = "#1a1a2e"
SUBTLE_FG = "#6c7086"
STEP_INACTIVE = "#45475a"
STEP_ACTIVE = "#89b4fa"
STEP_DONE = "#a6e3a1"
ROW_ALT_BG = "#282840"
DEFAULT_OUTPUT_DIR = str(Path("./generated_projects").resolve())


def _bind_hover(widget, normal_bg, hover_bg, normal_fg=None, hover_fg=None):
    """Add hover color effect to a widget."""
    def on_enter(e):
        widget.config(bg=hover_bg)
        if hover_fg:
            widget.config(fg=hover_fg)
    def on_leave(e):
        widget.config(bg=normal_bg)
        if normal_fg:
            widget.config(fg=normal_fg)
    widget.bind("<Enter>", on_enter)
    widget.bind("<Leave>", on_leave)


def _make_card(parent, pad_x=16, pad_y=12):
    """Create a card-style frame with a subtle border."""
    outer = tk.Frame(parent, bg=BORDER_COLOR, padx=1, pady=1)
    inner = tk.Frame(outer, bg=CARD_BG, padx=pad_x, pady=pad_y)
    inner.pack(fill="both", expand=True)
    outer._inner = inner
    return outer, inner


def _make_styled_entry(parent, var, width=50, show=None, font=("Consolas", 10)):
    """Create a styled entry with a subtle border frame."""
    border = tk.Frame(parent, bg=BORDER_COLOR, padx=1, pady=1)
    entry_opts = dict(
        textvariable=var, font=font,
        bg=ENTRY_BG, fg=ENTRY_FG, insertbackground=FG,
        relief="flat", width=width, bd=0,
    )
    if show:
        entry_opts["show"] = show
    ent = tk.Entry(border, **entry_opts)
    ent.pack(fill="x", ipady=5)
    return border, ent


# ──────────────────────────────────────────────────────────
# Main Application
# ──────────────────────────────────────────────────────────

class RegulationInitApp(tk.Tk):
    """Wizard-style Tkinter application."""

    STEP_LABELS = ["Code", "Type", "Git Name", "Reg. Files", "Out. Files", "GitLab", "Summary"]

    def __init__(self, manual_only=False):
        super().__init__()
        self.title("Regulation Project Initializer" + (" — Manual Entry" if manual_only else ""))
        self.geometry("880x750")
        self.configure(bg=BG)
        self.resizable(True, True)
        self.minsize(750, 600)
        self._manual_only = manual_only

        # Data collected across pages
        self.collected = {
            "regulation_code": "",
            "project_type": "both",
            "git_project_name": "",
            "output_base_dir": DEFAULT_OUTPUT_DIR,
            "regulation_file_map": {},
            "output_file_map": {},
            "regulation_identification": {},
            "dictionary_identification": {},
            "transformation_rules_identification": {},
            "transformed_data_identification": {},
            "ui_dashboard_identification": {},
            "output_identification": {},
            "output_structure_identification": {},
            "regulation_project_name": "",
            "output_project_name": "",
            "publish_gitlab": False,
            "gitlab_branch": "main",
            "gitlab_token": os.environ.get("GITLAB_TOKEN", ""),
        }

        # Page stack
        self._pages: list[tk.Frame] = []
        self._current_page_index = -1

        # Top accent bar
        tk.Frame(self, bg=ACCENT, height=3).pack(fill="x")

        # Container
        self._container = tk.Frame(self, bg=BG)
        self._container.pack(fill="both", expand=True, padx=22, pady=12)

        # Title bar
        title_frame = tk.Frame(self._container, bg=BG)
        title_frame.pack(fill="x", pady=(0, 4))
        tk.Label(
            title_frame, text="⚙  Regulation Project Initializer",
            font=("Segoe UI", 15, "bold"), bg=BG, fg=HEADER_FG,
        ).pack(anchor="w")

        # ── Visual step progress bar ──
        self._stepper_frame = tk.Frame(self._container, bg=BG)
        self._stepper_frame.pack(fill="x", pady=(8, 12))
        self._step_indicators = []
        self._step_labels_widgets = []
        self._step_connectors = []

        for i, label in enumerate(self.STEP_LABELS):
            if i > 0:
                connector = tk.Frame(self._stepper_frame, bg=STEP_INACTIVE, height=2, width=30)
                connector.pack(side="left", fill="x", expand=True, pady=(0, 14))
                self._step_connectors.append(connector)

            step_col = tk.Frame(self._stepper_frame, bg=BG)
            step_col.pack(side="left")

            circle = tk.Canvas(step_col, width=28, height=28, bg=BG, highlightthickness=0)
            circle.pack()
            oval_id = circle.create_oval(2, 2, 26, 26, fill=STEP_INACTIVE, outline="")
            text_id = circle.create_text(14, 14, text=str(i + 1), font=("Segoe UI", 8, "bold"), fill=FG)
            self._step_indicators.append((circle, oval_id, text_id))

            lbl = tk.Label(step_col, text=label, font=("Segoe UI", 7), bg=BG, fg=SUBTLE_FG)
            lbl.pack(pady=(2, 0))
            self._step_labels_widgets.append(lbl)

        # Separator below stepper
        ttk.Separator(self._container).pack(fill="x", pady=(0, 6))

        # Page area
        self._page_area = tk.Frame(self._container, bg=BG)
        self._page_area.pack(fill="both", expand=True)

        # Navigation bar
        nav = tk.Frame(self._container, bg=BG)
        nav.pack(fill="x", pady=(12, 0))

        ttk.Separator(nav).pack(fill="x", pady=(0, 10))

        nav_buttons = tk.Frame(nav, bg=BG)
        nav_buttons.pack(fill="x")

        self._btn_back = tk.Button(
            nav_buttons, text="◀  Back", font=("Segoe UI", 10),
            bg=ENTRY_BG, fg=FG, activebackground=BORDER_COLOR, activeforeground=FG,
            relief="flat", padx=16, pady=7, cursor="hand2",
            command=self._go_back,
        )
        self._btn_back.pack(side="left")
        _bind_hover(self._btn_back, ENTRY_BG, BORDER_COLOR, FG, FG)

        self._btn_next = tk.Button(
            nav_buttons, text="Next  ▶", font=("Segoe UI", 10, "bold"),
            bg=BTN_BG, fg=BTN_FG, activebackground=ACCENT_HOVER, activeforeground=BTN_FG,
            relief="flat", padx=16, pady=7, cursor="hand2",
            command=self._go_next,
        )
        self._btn_next.pack(side="right")
        _bind_hover(self._btn_next, BTN_BG, ACCENT_HOVER)

        self._progress_label = tk.Label(
            nav_buttons, text="", font=("Segoe UI", 9), bg=BG, fg=SUBTLE_FG,
        )
        self._progress_label.pack(side="right", padx=12)

        # Build pages
        self._build_pages()
        self._show_page(0)

    # ────────────── Page Factory ──────────────

    def _build_pages(self):
        self._pages = [
            self._page_regulation_code(),
            self._page_project_type(),
            self._page_git_name(),
            self._page_regulation_files(),
            self._page_output_files(),
            self._page_gitlab_push(),
            self._page_summary(),
        ]

    def _make_page(self) -> tk.Frame:
        frame = tk.Frame(self._page_area, bg=BG)
        return frame

    def _section_label(self, parent, text):
        frame = tk.Frame(parent, bg=BG)
        frame.pack(fill="x", pady=(14, 2))
        tk.Frame(frame, bg=ACCENT, width=4, height=20).pack(side="left", padx=(0, 10))
        tk.Label(
            frame, text=text, font=("Segoe UI", 13, "bold"),
            bg=BG, fg=ACCENT,
        ).pack(side="left")

    def _desc_label(self, parent, text):
        tk.Label(
            parent, text=text, font=("Segoe UI", 9),
            bg=BG, fg=SUBTLE_FG, wraplength=750, justify="left",
        ).pack(anchor="w", padx=(14, 0), pady=(0, 10))

    def _labeled_entry(self, parent, label, default="", width=50, show=None):
        row = tk.Frame(parent, bg=BG)
        row.pack(fill="x", pady=5, padx=(14, 0))
        tk.Label(row, text=label, font=("Segoe UI", 10), bg=BG, fg=FG, width=30, anchor="w").pack(side="left")
        var = tk.StringVar(value=default)
        border, _ = _make_styled_entry(row, var, width=width, show=show)
        border.pack(side="left", fill="x", expand=True, padx=(4, 0))
        return var

    def _labeled_check(self, parent, label, default=True):
        var = tk.BooleanVar(value=default)
        cb = tk.Checkbutton(
            parent, text=label, variable=var,
            font=("Segoe UI", 10), bg=BG, fg=FG,
            selectcolor=ENTRY_BG, activebackground=BG, activeforeground=FG,
            anchor="w",
        )
        cb.pack(fill="x", pady=2)
        return var

    # ────────────── Page 0: Regulation Code ──────────────

    def _page_regulation_code(self):
        page = self._make_page()
        self._section_label(page, "Regulatory Code")
        self._desc_label(page, "Enter the regulatory code (e.g.: COREP, FINREP, AnaCredit).")
        self._var_reg_code = self._labeled_entry(page, "Regulatory code", "")

        # ── Jira ticket import section (only shown when not in manual-only mode) ──
        if not self._manual_only:
            sep_frame = tk.Frame(page, bg=BG)
            sep_frame.pack(fill="x", pady=(18, 4))
            ttk.Separator(sep_frame).pack(fill="x")

            self._section_label(page, "Import from a Jira ticket")
            self._desc_label(
                page,
                "Select a file exported from Jira (.pdf, .doc, .html) to "
                "automatically pre-fill all project fields.",
            )

            # Card-style container for Jira import
            jira_card_border = tk.Frame(page, bg=BORDER_COLOR, padx=1, pady=1)
            jira_card_border.pack(fill="x", padx=(14, 0), pady=4)
            jira_card = tk.Frame(jira_card_border, bg=CARD_BG, padx=16, pady=12)
            jira_card.pack(fill="x")

            jira_btn_frame = tk.Frame(jira_card, bg=CARD_BG)
            jira_btn_frame.pack(fill="x")

            self._btn_import_jira = tk.Button(
                jira_btn_frame, text="📂  Import a Jira ticket",
                font=("Segoe UI", 10, "bold"),
                bg=ACCENT, fg=BTN_FG,
                activebackground=ACCENT_HOVER, activeforeground=BTN_FG,
                relief="flat", padx=16, pady=8, cursor="hand2",
                command=self._import_jira_ticket,
            )
            self._btn_import_jira.pack(side="left")
            _bind_hover(self._btn_import_jira, ACCENT, ACCENT_HOVER)

            self._btn_auto_generate = tk.Button(
                jira_btn_frame, text="🚀  Automatic Generation",
                font=("Segoe UI", 10, "bold"),
                bg=BTN_BG, fg=BTN_FG,
                activebackground=ACCENT_HOVER, activeforeground=BTN_FG,
                relief="flat", padx=16, pady=8, cursor="hand2",
                command=self._run_auto_generation,
            )
            # Hidden until a Jira ticket is imported
            self._btn_auto_generate.pack(side="left", padx=(10, 0))
            self._btn_auto_generate.pack_forget()

            self._jira_status_label = tk.Label(
                jira_btn_frame, text="", font=("Segoe UI", 9),
                bg=CARD_BG, fg=SUBTLE_FG,
            )
            self._jira_status_label.pack(side="left", padx=12)

            # Area to display parsed info
            info_border = tk.Frame(page, bg=BORDER_COLOR, padx=1, pady=1)
            info_border.pack(fill="both", expand=True, padx=(14, 0), pady=(8, 0))
            self._jira_info_text = scrolledtext.ScrolledText(
                info_border, font=("Consolas", 9), bg=ENTRY_BG, fg=FG,
                insertbackground=FG, relief="flat", height=10, wrap="word", bd=0,
            )
            self._jira_info_text.pack(fill="both", expand=True)
            self._jira_info_text.config(state="disabled")

        return page

    def _import_jira_ticket(self):
        """Open file dialog and parse a Jira ticket file (PDF, DOC, HTML)."""
        file_path = filedialog.askopenfilename(
            title="Select a Jira file",
            filetypes=[
                ("Jira files", "*.pdf *.doc *.html *.htm"),
                ("PDF", "*.pdf"),
                ("Word/HTML", "*.doc"),
                ("HTML", "*.html *.htm"),
                ("All files", "*.*"),
            ],
        )
        if not file_path:
            return

        try:
            jira_data = parse_jira_file(file_path)
            collected_data = jira_data_to_collected(jira_data)

            # Pre-fill regulation code
            if collected_data.get("regulation_code"):
                self._var_reg_code.set(collected_data["regulation_code"])

            # Pre-fill project type
            if collected_data.get("project_type"):
                self._var_project_type.set(collected_data["project_type"])

            # Pre-fill git name
            if collected_data.get("git_project_name"):
                self._var_git_name.set(collected_data["git_project_name"])

            # Pre-fill identification data in collected dict
            for key in [
                "regulation_identification", "dictionary_identification",
                "transformation_rules_identification", "transformed_data_identification",
                "ui_dashboard_identification", "output_identification",
                "output_structure_identification",
            ]:
                if collected_data.get(key):
                    self.collected[key] = collected_data[key]

            # Store extra Jira info
            self.collected["jira_ticket_key"] = collected_data.get("jira_ticket_key", "")
            self.collected["jira_title"] = collected_data.get("jira_title", "")
            self.collected["regulation_project_name"] = collected_data.get("regulation_project_name", "")
            self.collected["output_project_name"] = collected_data.get("output_project_name", "")
            self.collected["git_project_url"] = collected_data.get("git_project_url", "")

            # Update file name defaults with the parsed regulation code
            code = collected_data.get("regulation_code", "")
            if code:
                jira_mappings = collected_data.get("file_mappings", {})
                for check_var, ref_name, name_var in self._reg_file_vars:
                    resolved = self._resolve_jira_file_name(ref_name, code, jira_mappings)
                    name_var.set(resolved)
                for check_var, ref_name, name_var in self._out_file_vars:
                    resolved = self._resolve_jira_file_name(ref_name, code, jira_mappings)
                    name_var.set(resolved)

            # Store collected data for auto mode
            self._jira_collected = collected_data

            # Show summary
            summary = format_jira_summary(jira_data)
            self._jira_info_text.config(state="normal")
            self._jira_info_text.delete("1.0", "end")
            self._jira_info_text.insert("end", summary)
            self._jira_info_text.config(state="disabled")

            self._jira_status_label.config(
                text=f"✅ Ticket {jira_data.ticket_key} imported",
                fg=SUCCESS_FG,
            )

            # Show auto-generate button
            self._btn_auto_generate.pack(side="left", padx=(10, 0))

            logger.info("Jira ticket imported: %s", file_path)

        except Exception as exc:
            self._jira_status_label.config(
                text=f"❌ Error: {exc}",
                fg=ERR_FG,
            )
            self._jira_info_text.config(state="normal")
            self._jira_info_text.delete("1.0", "end")
            self._jira_info_text.insert("end", f"Parsing error:\n{exc}")
            self._jira_info_text.config(state="disabled")
            logger.error("Failed to parse Jira ticket: %s", exc)

    @staticmethod
    def _resolve_jira_file_name(ref_name: str, regulation_code: str, jira_mappings: dict) -> str:
        """Resolve target file name: use Jira explicit mapping if available, else derive from code."""
        # Try exact match first
        if ref_name in jira_mappings:
            return jira_mappings[ref_name]
        # Try case-insensitive exact match
        for jira_src, jira_dst in jira_mappings.items():
            if jira_src.lower() == ref_name.lower():
                return jira_dst
        # Try prefix match (everything before the last '-')
        for jira_src, jira_dst in jira_mappings.items():
            jira_prefix = jira_src.rsplit("-", 1)[0] if "-" in jira_src else jira_src
            ref_prefix = ref_name.rsplit("-", 1)[0] if "-" in ref_name else ref_name
            if jira_prefix.lower() == ref_prefix.lower():
                return jira_dst
        return _tokenized_default_name(ref_name, regulation_code)

    def _run_auto_generation(self):
        """Auto-generate everything from imported Jira data without manual steps."""
        collected = getattr(self, "_jira_collected", None)
        if not collected:
            messagebox.showwarning("No ticket", "Import a Jira ticket first.")
            return

        confirm = messagebox.askyesno(
            "Automatic Generation",
            "All files will be automatically generated from the Jira ticket.\n\n"
            f"Code: {collected.get('regulation_code', '?')}\n"
            f"Type: {collected.get('project_type', '?')}\n"
            f"Project: {collected.get('git_project_name', '?')}\n\n"
            "Continue?",
        )
        if not confirm:
            return

        code = collected.get("regulation_code", "")
        project_type = collected.get("project_type", "both")
        git_name = _sanitize_project_name(collected.get("git_project_name", "") or code.lower())
        jira_mappings = collected.get("file_mappings", {})

        # Build file maps for all reference files
        reg_file_map = {}
        out_file_map = {}

        if project_type in ("both", "regulation"):
            ref_dir = _reference_dir_for_scope("regulation")
            if ref_dir.exists():
                for f in sorted(ref_dir.iterdir()):
                    if f.is_file():
                        target = self._resolve_jira_file_name(f.name, code, jira_mappings)
                        final = _sanitize_file_name(target, f.suffix)
                        if final:
                            reg_file_map[f.name] = final

        if project_type in ("both", "output"):
            ref_dir = _reference_dir_for_scope("output")
            if ref_dir.exists():
                for f in sorted(ref_dir.iterdir()):
                    if f.is_file():
                        target = self._resolve_jira_file_name(f.name, code, jira_mappings)
                        final = _sanitize_file_name(target, f.suffix)
                        if final:
                            out_file_map[f.name] = final

        # Store collected data
        self.collected.update({
            "regulation_code": code,
            "project_type": project_type,
            "git_project_name": git_name,
            "regulation_file_map": reg_file_map,
            "output_file_map": out_file_map,
            "publish_gitlab": False,
            "gitlab_branch": "main",
        })

        # Jump to summary page and auto-run
        self._show_page(len(self._pages) - 1)
        self._populate_summary()
        self._run_generation()

    # ────────────── Page 1: Project Type ──────────────

    def _page_project_type(self):
        page = self._make_page()
        self._section_label(page, "Project Type")
        self._desc_label(page, "Choose the type of project to create.")

        self._var_project_type = tk.StringVar(value="both")
        options = [
            ("📦  Regulation + Output (Both)", "both", "Generate both regulation and output files"),
            ("📋  Regulation only", "regulation", "Generate only regulation files"),
            ("📊  Output only", "output", "Generate only output files"),
        ]
        for text, val, desc in options:
            opt_border = tk.Frame(page, bg=BORDER_COLOR, padx=1, pady=1)
            opt_border.pack(fill="x", pady=4, padx=(14, 0))
            opt_frame = tk.Frame(opt_border, bg=CARD_BG, padx=14, pady=10)
            opt_frame.pack(fill="x")
            rb = tk.Radiobutton(
                opt_frame, text=text, variable=self._var_project_type, value=val,
                font=("Segoe UI", 10, "bold"), bg=CARD_BG, fg=FG,
                selectcolor=ENTRY_BG, activebackground=CARD_BG, activeforeground=FG,
                anchor="w",
            )
            rb.pack(fill="x")
            tk.Label(
                opt_frame, text=desc, font=("Segoe UI", 8), bg=CARD_BG, fg=SUBTLE_FG,
                anchor="w",
            ).pack(fill="x", padx=(24, 0))

        return page

    # ────────────── Page 2: Git Project Name ──────────────

    def _page_git_name(self):
        page = self._make_page()
        self._section_label(page, "Git Project Name")
        self._desc_label(
            page,
            "Base name for the Git project. Example: corep → <destination>/corep/regulation and/or output",
        )
        self._var_git_name = self._labeled_entry(page, "Git project name", "")

        # ── Destination folder chooser ──
        self._section_label(page, "Destination Folder")
        self._desc_label(page, "Choose the folder where the project will be created.")

        dest_row = tk.Frame(page, bg=BG)
        dest_row.pack(fill="x", pady=5, padx=(14, 0))
        tk.Label(dest_row, text="Destination folder", font=("Segoe UI", 10), bg=BG, fg=FG, width=30, anchor="w").pack(side="left")

        self._var_output_dir = tk.StringVar(value=DEFAULT_OUTPUT_DIR)
        border, self._ent_output_dir = _make_styled_entry(dest_row, self._var_output_dir, width=40)
        border.pack(side="left", fill="x", expand=True, padx=(4, 0))

        btn_browse = tk.Button(
            dest_row, text="\U0001f4c2 Browse", font=("Segoe UI", 9),
            bg=BTN_BG, fg=BTN_FG, activebackground=ACCENT_HOVER, activeforeground=BTN_FG,
            relief="flat", padx=10, pady=4, cursor="hand2",
            command=self._browse_output_dir,
        )
        btn_browse.pack(side="left", padx=(6, 0))
        _bind_hover(btn_browse, BTN_BG, ACCENT_HOVER)

        return page

    def _browse_output_dir(self):
        """Open a folder chooser for the project destination."""
        current = self._var_output_dir.get().strip()
        initial = current if current and Path(current).exists() else str(Path.home())
        chosen = filedialog.askdirectory(title="Select destination folder", initialdir=initial)
        if chosen:
            self._var_output_dir.set(chosen)

    # ────────────── Page 3: Regulation Files ──────────────

    def _page_regulation_files(self):
        page = self._make_page()
        self._section_label(page, "Regulation Files")
        self._desc_label(page, "Check the files to generate and modify the target name if needed.")

        # Header row
        header = tk.Frame(page, bg=SECTION_BG, padx=14, pady=6)
        header.pack(fill="x", padx=(14, 0), pady=(0, 2))
        tk.Label(header, text="", width=3, bg=SECTION_BG).pack(side="left")
        tk.Label(header, text="Source file", font=("Segoe UI", 9, "bold"), bg=SECTION_BG, fg=SUBTLE_FG, width=42, anchor="w").pack(side="left")
        tk.Label(header, text="→", font=("Segoe UI", 9), bg=SECTION_BG, fg=SUBTLE_FG).pack(side="left", padx=4)
        tk.Label(header, text="Target name", font=("Segoe UI", 9, "bold"), bg=SECTION_BG, fg=SUBTLE_FG, anchor="w").pack(side="left")

        self._reg_file_vars: list[tuple[tk.BooleanVar, str, tk.StringVar]] = []
        ref_dir = _reference_dir_for_scope("regulation")
        if ref_dir.exists():
            files = sorted([p for p in ref_dir.iterdir() if p.is_file()])
            for i, f in enumerate(files):
                row_bg = ROW_ALT_BG if i % 2 == 0 else BG
                row = tk.Frame(page, bg=row_bg, padx=14, pady=5)
                row.pack(fill="x", padx=(14, 0))

                check_var = tk.BooleanVar(value=True)
                cb = tk.Checkbutton(
                    row, variable=check_var, bg=row_bg, selectcolor=ENTRY_BG,
                    activebackground=row_bg,
                )
                cb.pack(side="left")

                tk.Label(
                    row, text=f.name, font=("Consolas", 9), bg=row_bg, fg=FG, width=42, anchor="w",
                ).pack(side="left")

                tk.Label(row, text="→", font=("Segoe UI", 9), bg=row_bg, fg=ACCENT).pack(side="left", padx=6)

                name_var = tk.StringVar(value=f.name)
                ent = tk.Entry(
                    row, textvariable=name_var, font=("Consolas", 9),
                    bg=ENTRY_BG, fg=ENTRY_FG, insertbackground=FG, relief="flat", width=42,
                )
                ent.pack(side="left", ipady=4, padx=(4, 0))
                self._reg_file_vars.append((check_var, f.name, name_var))

        return page

    # ────────────── Page 4: Output Files ──────────────

    def _page_output_files(self):
        page = self._make_page()
        self._section_label(page, "Output Files")
        self._desc_label(page, "Check the files to generate and modify the target name if needed.")

        # Header row
        header = tk.Frame(page, bg=SECTION_BG, padx=14, pady=6)
        header.pack(fill="x", padx=(14, 0), pady=(0, 2))
        tk.Label(header, text="", width=3, bg=SECTION_BG).pack(side="left")
        tk.Label(header, text="Source file", font=("Segoe UI", 9, "bold"), bg=SECTION_BG, fg=SUBTLE_FG, width=50, anchor="w").pack(side="left")
        tk.Label(header, text="→", font=("Segoe UI", 9), bg=SECTION_BG, fg=SUBTLE_FG).pack(side="left", padx=4)
        tk.Label(header, text="Target name", font=("Segoe UI", 9, "bold"), bg=SECTION_BG, fg=SUBTLE_FG, anchor="w").pack(side="left")

        self._out_file_vars: list[tuple[tk.BooleanVar, str, tk.StringVar]] = []
        ref_dir = _reference_dir_for_scope("output")
        if ref_dir.exists():
            files = sorted([p for p in ref_dir.iterdir() if p.is_file()])
            for i, f in enumerate(files):
                row_bg = ROW_ALT_BG if i % 2 == 0 else BG
                row = tk.Frame(page, bg=row_bg, padx=14, pady=5)
                row.pack(fill="x", padx=(14, 0))

                check_var = tk.BooleanVar(value=True)
                cb = tk.Checkbutton(
                    row, variable=check_var, bg=row_bg, selectcolor=ENTRY_BG,
                    activebackground=row_bg,
                )
                cb.pack(side="left")

                tk.Label(
                    row, text=f.name, font=("Consolas", 9), bg=row_bg, fg=FG, width=50, anchor="w",
                ).pack(side="left")

                tk.Label(row, text="→", font=("Segoe UI", 9), bg=row_bg, fg=ACCENT).pack(side="left", padx=6)

                name_var = tk.StringVar(value=f.name)
                ent = tk.Entry(
                    row, textvariable=name_var, font=("Consolas", 9),
                    bg=ENTRY_BG, fg=ENTRY_FG, insertbackground=FG, relief="flat", width=36,
                )
                ent.pack(side="left", ipady=4, padx=(4, 0))
                self._out_file_vars.append((check_var, f.name, name_var))

        return page

    # ────────────── Page 5: GitLab Push ──────────────

    def _page_gitlab_push(self):
        page = self._make_page()
        self._section_label(page, "GitLab Publishing")
        self._desc_label(page, "Configure publishing to GitLab (optional).")

        # Card for the checkbox
        check_card_border = tk.Frame(page, bg=BORDER_COLOR, padx=1, pady=1)
        check_card_border.pack(fill="x", padx=(14, 0), pady=(0, 12))
        check_card = tk.Frame(check_card_border, bg=CARD_BG, padx=14, pady=10)
        check_card.pack(fill="x")

        self._var_publish = tk.BooleanVar(value=False)
        tk.Checkbutton(
            check_card, text="  Push files to GitLab after generation",
            variable=self._var_publish,
            font=("Segoe UI", 10, "bold"), bg=CARD_BG, fg=FG,
            selectcolor=ENTRY_BG, activebackground=CARD_BG, activeforeground=FG,
        ).pack(fill="x")

        self._var_branch = self._labeled_entry(page, "Target branch", "main")
        self._var_token = self._labeled_entry(page, "GitLab Token (PRIVATE-TOKEN)", os.environ.get("GITLAB_TOKEN", ""), show="●")

        return page

    # ────────────── Page 6: Summary + Run ──────────────

    def _page_summary(self):
        page = self._make_page()
        self._section_label(page, "Summary & Generation")
        self._desc_label(page, "Review your configuration before generating files.")

        summary_border = tk.Frame(page, bg=BORDER_COLOR, padx=1, pady=1)
        summary_border.pack(fill="both", expand=True, padx=(14, 0), pady=6)
        self._summary_text = scrolledtext.ScrolledText(
            summary_border, font=("Consolas", 9), bg=SECTION_BG, fg=FG,
            insertbackground=FG, relief="flat", height=20, wrap="word", bd=0,
        )
        self._summary_text.pack(fill="both", expand=True)

        # Configure text tags for colors
        self._summary_text.tag_config("header", foreground=HEADER_FG, font=("Consolas", 10, "bold"))
        self._summary_text.tag_config("label", foreground=ACCENT)
        self._summary_text.tag_config("value", foreground=FG, font=("Consolas", 9, "bold"))
        self._summary_text.tag_config("success", foreground=SUCCESS_FG)
        self._summary_text.tag_config("warn", foreground=WARN_FG)
        self._summary_text.tag_config("file_arrow", foreground=ACCENT)

        return page

    # ────────────── Navigation ──────────────

    def _show_page(self, idx):
        if 0 <= self._current_page_index < len(self._pages):
            self._pages[self._current_page_index].pack_forget()

        self._current_page_index = idx
        self._pages[idx].pack(fill="both", expand=True)

        # Update nav buttons
        self._btn_back.config(state="normal" if idx > 0 else "disabled")
        is_last = idx == len(self._pages) - 1
        self._btn_next.config(text="🚀  Generate" if is_last else "Next  ▶")
        self._progress_label.config(text=f"Step {idx + 1} / {len(self._pages)}")

        # Update stepper visuals
        for i, (canvas, oval_id, text_id) in enumerate(self._step_indicators):
            if i < idx:
                canvas.itemconfig(oval_id, fill=STEP_DONE)
                canvas.itemconfig(text_id, text="✓", fill=BG)
                self._step_labels_widgets[i].config(fg=STEP_DONE)
            elif i == idx:
                canvas.itemconfig(oval_id, fill=STEP_ACTIVE)
                canvas.itemconfig(text_id, text=str(i + 1), fill=BG)
                self._step_labels_widgets[i].config(fg=STEP_ACTIVE, font=("Segoe UI", 7, "bold"))
            else:
                canvas.itemconfig(oval_id, fill=STEP_INACTIVE)
                canvas.itemconfig(text_id, text=str(i + 1), fill=FG)
                self._step_labels_widgets[i].config(fg=SUBTLE_FG, font=("Segoe UI", 7))

        for i, connector in enumerate(self._step_connectors):
            connector.config(bg=STEP_DONE if i < idx else STEP_INACTIVE)

        # Populate summary on last page
        if is_last:
            self._populate_summary()

    def _go_back(self):
        if self._current_page_index > 0:
            target = self._prev_visible_page(self._current_page_index)
            if target is not None:
                self._show_page(target)

    def _go_next(self):
        if not self._validate_current_page():
            return

        # Collect current page data
        self._collect_current_page()

        if self._current_page_index == len(self._pages) - 1:
            self._run_generation()
            return

        target = self._next_visible_page(self._current_page_index)
        if target is not None:
            self._show_page(target)

    def _next_visible_page(self, current):
        """Skip pages that don't apply based on project type."""
        idx = current + 1
        while idx < len(self._pages):
            if self._page_is_visible(idx):
                return idx
            idx += 1
        return None

    def _prev_visible_page(self, current):
        idx = current - 1
        while idx >= 0:
            if self._page_is_visible(idx):
                return idx
            idx -= 1
        return None

    def _page_is_visible(self, idx):
        ptype = self._var_project_type.get()
        # Page 3 = regulation files → skip if output only
        if idx == 3 and ptype == "output":
            return False
        # Page 4 = output files → skip if regulation only
        if idx == 4 and ptype == "regulation":
            return False
        return True

    # ────────────── Validation ──────────────

    def _validate_current_page(self):
        idx = self._current_page_index

        if idx == 0:
            code = self._var_reg_code.get().strip()
            if not code:
                messagebox.showwarning("Required field", "The regulatory code is required.")
                return False

        if idx == 2:
            code = self._var_reg_code.get().strip().upper()
            raw = self._var_git_name.get().strip()
            candidate = _sanitize_project_name(raw or code.lower())
            if not candidate:
                messagebox.showwarning("Invalid name", "Use letters/digits/hyphens.")
                return False

        return True

    # ────────────── Data Collection ──────────────

    def _collect_current_page(self):
        idx = self._current_page_index
        d = self.collected

        if idx == 0:
            d["regulation_code"] = self._var_reg_code.get().strip().upper()
            # Update default file names with the regulation code
            code = d["regulation_code"]
            for _, ref_name, name_var in self._reg_file_vars:
                name_var.set(_tokenized_default_name(ref_name, code))
            for _, ref_name, name_var in self._out_file_vars:
                name_var.set(_tokenized_default_name(ref_name, code))
            # Update default git name
            if not self._var_git_name.get().strip():
                self._var_git_name.set(_sanitize_project_name(code.lower()))

        elif idx == 1:
            d["project_type"] = self._var_project_type.get()

        elif idx == 2:
            code = d["regulation_code"]
            raw = self._var_git_name.get().strip()
            d["git_project_name"] = _sanitize_project_name(raw or code.lower())
            d["output_base_dir"] = self._var_output_dir.get().strip() or DEFAULT_OUTPUT_DIR

        elif idx == 3:
            fmap = {}
            for check_var, ref_name, name_var in self._reg_file_vars:
                if check_var.get():
                    final = _sanitize_file_name(name_var.get(), Path(ref_name).suffix)
                    if final:
                        fmap[ref_name] = final
            d["regulation_file_map"] = fmap

        elif idx == 4:
            fmap = {}
            for check_var, ref_name, name_var in self._out_file_vars:
                if check_var.get():
                    final = _sanitize_file_name(name_var.get(), Path(ref_name).suffix)
                    if final:
                        fmap[ref_name] = final
            d["output_file_map"] = fmap

        elif idx == 5:
            d["publish_gitlab"] = self._var_publish.get()
            d["gitlab_branch"] = _sanitize_branch_name(self._var_branch.get().strip()) or "main"
            d["gitlab_token"] = self._var_token.get().strip()

    # ────────────── Summary ──────────────

    def _populate_summary(self):
        self._collect_current_page()
        d = self.collected
        txt = self._summary_text
        txt.config(state="normal")
        txt.delete("1.0", "end")

        txt.insert("end", "\n")
        txt.insert("end", "  ╔═══════════════════════════════════════════╗\n", "header")
        txt.insert("end", "  ║       CONFIGURATION SUMMARY              ║\n", "header")
        txt.insert("end", "  ╚═══════════════════════════════════════════╝\n", "header")
        txt.insert("end", "\n")

        txt.insert("end", "  Regulatory code    :  ", "label")
        txt.insert("end", f"{d['regulation_code']}\n", "value")
        txt.insert("end", "  Project type       :  ", "label")
        txt.insert("end", f"{d['project_type']}\n", "value")
        txt.insert("end", "  Git project name   :  ", "label")
        txt.insert("end", f"{d['git_project_name']}\n", "value")
        txt.insert("end", "  Destination folder :  ", "label")
        txt.insert("end", f"{d['output_base_dir']}\n\n", "value")

        if d["project_type"] in ("both", "regulation"):
            txt.insert("end", "  ── Regulation Files ──\n", "header")
            for ref, target in d["regulation_file_map"].items():
                txt.insert("end", "    ✓ ", "success")
                txt.insert("end", f"{ref}", "label")
                txt.insert("end", "  →  ", "file_arrow")
                txt.insert("end", f"{target}\n", "value")
            if not d["regulation_file_map"]:
                txt.insert("end", "    (no file selected)\n", "warn")
            txt.insert("end", "\n")

        if d["project_type"] in ("both", "output"):
            txt.insert("end", "  ── Output Files ──\n", "header")
            for ref, target in d["output_file_map"].items():
                txt.insert("end", "    ✓ ", "success")
                txt.insert("end", f"{ref}", "label")
                txt.insert("end", "  →  ", "file_arrow")
                txt.insert("end", f"{target}\n", "value")
            if not d["output_file_map"]:
                txt.insert("end", "    (no file selected)\n", "warn")
            txt.insert("end", "\n")

        if d["publish_gitlab"]:
            txt.insert("end", "  GitLab Push        :  ", "label")
            txt.insert("end", f"YES  (branch: {d['gitlab_branch']})\n", "success")
            txt.insert("end", "  Token              :  ", "label")
            txt.insert("end", f"{'***' if d['gitlab_token'] else 'NOT CONFIGURED'}\n", "value")
        else:
            txt.insert("end", "  GitLab Push        :  ", "label")
            txt.insert("end", "NO (local generation only)\n", "warn")

        txt.insert("end", "\n")
        txt.insert("end", "  ═══════════════════════════════════════════\n", "header")
        txt.insert("end", "  Click « 🚀 Generate » to start.\n", "label")
        txt.config(state="disabled")

    # ────────────── Generation ──────────────

    def _run_generation(self):
        self._collect_current_page()
        d = self.collected

        # Disable button during generation
        self._btn_next.config(state="disabled", text="⏳  Generation in progress…")

        # Unlock the summary text to show progress
        self._summary_text.config(state="normal")
        self._summary_text.insert("end", "\n\n▶ Starting generation…\n")
        self._summary_text.see("end")
        self.update_idletasks()

        def _do_generate():
            try:
                token = d["gitlab_token"]
                config = AppConfig(
                    gitlab=GitLabConfig(private_token=token),
                    output_base_dir=Path(d["output_base_dir"]),
                    project_base_name=d["git_project_name"],
                    project_type=d["project_type"],
                    regulation_folder_name="",
                    output_folder_name="",
                    regulation_file_map=d["regulation_file_map"],
                    output_file_map=d["output_file_map"],
                    dictionary_identification=d["dictionary_identification"],
                    regulation_identification=d["regulation_identification"],
                    output_identification=d["output_identification"],
                    transformation_rules_identification=d["transformation_rules_identification"],
                    transformed_data_identification=d["transformed_data_identification"],
                    ui_dashboard_identification=d["ui_dashboard_identification"],
                    output_structure_identification=d["output_structure_identification"],
                    regulation_project_name=d.get("regulation_project_name", ""),
                    output_project_name=d.get("output_project_name", ""),
                )

                regulation = RegulationInfo(
                    code=d["regulation_code"], version="", name="", authority="",
                    frequency="", first_reporting_date="",
                )
                project_input = ProjectInput(regulation=regulation)
                initializer = ProjectInitializer(config, project_input)
                result_dir = initializer.run()

                # GitLab publish
                if d.get("publish_gitlab") and token:
                    self.after(0, lambda: self._log_progress("▶ Publishing to GitLab…\n"))
                    publisher = GitLabPublisher(config.gitlab)
                    reg_code = d["regulation_code"]
                    username = publisher.get_authenticated_username()
                    project_base = d["git_project_name"]
                    branch = d.get("gitlab_branch", "main")

                    # Use the namespace from Jira git_project_url if available
                    git_namespace = _extract_gitlab_namespace(d.get("git_project_url", ""))

                    create_regulation = d["project_type"] in ("both", "regulation")
                    create_output = d["project_type"] in ("both", "output")

                    reg_folder = d.get("regulation_project_name") or f"pure-regulation-{project_base}"
                    out_folder = d.get("output_project_name") or f"pure-regulatoryoutput-{project_base}"

                    if create_regulation:
                        reg_project = f"{git_namespace}/{reg_folder}" if git_namespace else f"{username}/{reg_folder}"
                        publisher.ensure_project_exists(
                            project_path_with_namespace=reg_project,
                            username=username,
                            create_if_missing=True,
                        )
                        publisher.publish_directory(
                            project_path_with_namespace=reg_project,
                            local_dir=result_dir / reg_folder,
                            branch=branch,
                            commit_prefix=f"chore: init {reg_code} regulation",
                        )

                    if create_output:
                        out_project = f"{git_namespace}/{out_folder}" if git_namespace else f"{username}/{out_folder}"
                        publisher.ensure_project_exists(
                            project_path_with_namespace=out_project,
                            username=username,
                            create_if_missing=True,
                        )
                        publisher.publish_directory(
                            project_path_with_namespace=out_project,
                            local_dir=result_dir / out_folder,
                            branch=branch,
                            commit_prefix=f"chore: init {reg_code} output",
                        )

                # Build tree string
                tree_lines = []
                self._build_tree_lines(result_dir, tree_lines)

                self.after(0, lambda: self._on_generation_success(result_dir, tree_lines))

            except Exception as exc:
                self.after(0, lambda e=exc: self._on_generation_error(e))

        threading.Thread(target=_do_generate, daemon=True).start()

    def _on_generation_success(self, result_dir, tree_lines):
        txt = self._summary_text
        txt.config(state="normal")
        txt.insert("end", "\n✅  Generation completed successfully!\n\n")
        txt.insert("end", f"📂  Files generated in: {result_dir.resolve()}\n\n")
        for line in tree_lines:
            txt.insert("end", line + "\n")
        txt.see("end")
        txt.config(state="disabled")

        self._btn_next.config(state="normal", text="✅  Done")
        self._btn_next.config(command=self.destroy)

    def _log_progress(self, message):
        txt = self._summary_text
        txt.config(state="normal")
        txt.insert("end", message)
        txt.see("end")
        txt.config(state="disabled")
        self.update_idletasks()

    def _on_generation_error(self, exc):
        txt = self._summary_text
        txt.config(state="normal")
        txt.insert("end", f"\n❌  Error: {exc}\n")
        txt.see("end")
        txt.config(state="disabled")

        self._btn_next.config(state="normal", text="🚀  Retry")
        self._btn_next.config(command=self._run_generation)

    def _build_tree_lines(self, directory: Path, lines: list, prefix: str = ""):
        entries = sorted(directory.iterdir(), key=lambda e: (e.is_file(), e.name))
        for i, entry in enumerate(entries):
            is_last = i == len(entries) - 1
            connector = "└── " if is_last else "├── "
            lines.append(f"{prefix}{connector}{entry.name}")
            if entry.is_dir():
                extension = "    " if is_last else "│   "
                self._build_tree_lines(entry, lines, prefix + extension)


# ──────────────────────────────────────────────────────────
# Identification questionnaire dialogs
# ──────────────────────────────────────────────────────────

class IdentificationDialog(tk.Toplevel):
    """Reusable modal dialog for collecting identification fields."""

    def __init__(self, parent, title, fields, auto_translate=None):
        """
        fields: list of (label, required, default)
        auto_translate: dict of {target_field: (source_field, fn)}
        """
        super().__init__(parent)
        self.title(title)
        self.configure(bg=BG)
        self.geometry("600x550")
        self.resizable(False, True)
        self.transient(parent)
        self.grab_set()

        self.result = None
        self._vars: dict[str, tk.StringVar] = {}
        self._fields = fields
        self._auto_translate = auto_translate or {}

        # Center on parent
        self.update_idletasks()
        px = parent.winfo_rootx() + (parent.winfo_width() - 600) // 2
        py = parent.winfo_rooty() + (parent.winfo_height() - 550) // 2
        self.geometry(f"600x550+{px}+{py}")

        # Top accent bar
        tk.Frame(self, bg=HEADER_FG, height=3).pack(fill="x")

        # Title
        title_frame = tk.Frame(self, bg=BG)
        title_frame.pack(fill="x", padx=20, pady=(16, 4))
        tk.Frame(title_frame, bg=HEADER_FG, width=4, height=20).pack(side="left", padx=(0, 10))
        tk.Label(
            title_frame, text=title, font=("Segoe UI", 13, "bold"), bg=BG, fg=HEADER_FG,
        ).pack(side="left")

        tk.Label(
            self, text="Fields marked with * are required",
            font=("Segoe UI", 8), bg=BG, fg=SUBTLE_FG,
        ).pack(anchor="w", padx=20, pady=(0, 8))

        ttk.Separator(self).pack(fill="x", padx=20)

        # Scrollable form
        canvas = tk.Canvas(self, bg=BG, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        form = tk.Frame(canvas, bg=BG)

        form.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=form, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True, padx=(20, 0), pady=10)
        scrollbar.pack(side="right", fill="y", padx=(0, 4), pady=10)

        for i, (label, required, default) in enumerate(fields):
            row_bg = ROW_ALT_BG if i % 2 == 0 else BG
            row = tk.Frame(form, bg=row_bg, padx=10, pady=8)
            row.pack(fill="x")

            marker = " *" if required else ""
            lbl_fg = ACCENT if required else FG
            tk.Label(
                row, text=f"{label}{marker}", font=("Segoe UI", 9, "bold" if required else ""),
                bg=row_bg, fg=lbl_fg, width=28, anchor="w",
            ).pack(side="left")

            var = tk.StringVar(value=default)
            border_frame = tk.Frame(row, bg=BORDER_COLOR, padx=1, pady=1)
            border_frame.pack(side="left", fill="x", expand=True, padx=(4, 0))
            ent = tk.Entry(
                border_frame, textvariable=var, font=("Consolas", 9),
                bg=ENTRY_BG, fg=ENTRY_FG, insertbackground=FG, relief="flat", width=30, bd=0,
            )
            ent.pack(fill="x", ipady=4)
            self._vars[label] = var

        # Button bar
        btn_frame = tk.Frame(self, bg=BG)
        btn_frame.pack(fill="x", padx=20, pady=14)
        ttk.Separator(btn_frame).pack(fill="x", pady=(0, 12))

        btn_row = tk.Frame(btn_frame, bg=BG)
        btn_row.pack(fill="x")

        btn_submit = tk.Button(
            btn_row, text="✓  Submit", font=("Segoe UI", 10, "bold"),
            bg=BTN_BG, fg=BTN_FG, relief="flat", padx=20, pady=6, cursor="hand2",
            command=self._submit,
        )
        btn_submit.pack(side="right")
        _bind_hover(btn_submit, BTN_BG, ACCENT_HOVER)

        btn_cancel = tk.Button(
            btn_row, text="Cancel", font=("Segoe UI", 10),
            bg=ENTRY_BG, fg=FG, relief="flat", padx=20, pady=6, cursor="hand2",
            command=self.destroy,
        )
        btn_cancel.pack(side="right", padx=(0, 8))
        _bind_hover(btn_cancel, ENTRY_BG, BORDER_COLOR, FG, FG)

    def _submit(self):
        result = {}
        for label, required, _ in self._fields:
            val = self._vars[label].get().strip()
            if required and not val:
                messagebox.showwarning("Required field", f"The field « {label} » is required.", parent=self)
                return
            result[label] = val

        # Apply auto-translations
        for target, (source, fn) in self._auto_translate.items():
            if source in result and target in result and not result[target]:
                result[target] = fn(result[source])

        self.result = result
        self.destroy()


# ──────────────────────────────────────────────────────────
# Extended App with identification dialogs
# ──────────────────────────────────────────────────────────

class RegulationInitAppFull(RegulationInitApp):
    """Extended app that opens identification dialogs between file selection and summary."""

    def __init__(self, manual_only=False):
        super().__init__(manual_only=manual_only)

    def _go_next(self):
        if not self._validate_current_page():
            return
        self._collect_current_page()

        idx = self._current_page_index

        # After regulation files page (3) → open identification dialogs
        if idx == 3:
            self._open_regulation_identification_dialogs()

        # After output files page (4) → open output identification dialogs
        if idx == 4:
            self._open_output_identification_dialogs()

        if self._current_page_index == len(self._pages) - 1:
            self._run_generation()
            return

        target = self._next_visible_page(self._current_page_index)
        if target is not None:
            self._show_page(target)

    def _open_regulation_identification_dialogs(self):
        d = self.collected
        fmap = d["regulation_file_map"]
        reg_id = d.get("regulation_identification", {})
        dict_id = d.get("dictionary_identification", {})
        tr_id = d.get("transformation_rules_identification", {})
        td_id = d.get("transformed_data_identification", {})
        ui_id = d.get("ui_dashboard_identification", {})

        if "pure-regulation-regulationCode.xlsx" in fmap:
            dlg = IdentificationDialog(
                self, "Regulation – Identification",
                [
                    ("regulatorCode", True, reg_id.get("regulatorCode", "")),
                    ("version", True, reg_id.get("version", "")),
                    ("validityStartingDate", True, reg_id.get("validityStartingDate", "")),
                    ("name (FR)", True, reg_id.get("name (FR)", "")),
                    ("name (EN)", False, reg_id.get("name (EN)", "")),
                    ("description (EN)", False, reg_id.get("description (EN)", "")),
                    ("description (FR)", False, reg_id.get("description (FR)", "")),
                    ("author", True, reg_id.get("author", "")),
                ],
                auto_translate={
                    "name (EN)": ("name (FR)", _translate_fr_to_en),
                    "description (EN)": ("name (EN)", lambda x: x),
                    "description (FR)": ("name (FR)", lambda x: x),
                },
            )
            self.wait_window(dlg)
            if dlg.result:
                d["regulation_identification"] = dlg.result

        if "pure-dictionary-regulatorCode.xlsx" in fmap:
            dlg = IdentificationDialog(
                self, "Dictionary – Identification",
                [
                    ("version", True, dict_id.get("version", "")),
                    ("validityStartingDate", True, dict_id.get("validityStartingDate", "")),
                    ("description (EN)", True, dict_id.get("description (EN)", "")),
                    ("description (FR)", False, dict_id.get("description (FR)", "")),
                    ("regulatorName (FR)", True, dict_id.get("regulatorName (FR)", "")),
                    ("regulatorName (EN)", False, dict_id.get("regulatorName (EN)", "")),
                ],
                auto_translate={
                    "description (FR)": ("description (EN)", _translate_en_to_fr),
                    "regulatorName (EN)": ("regulatorName (FR)", _translate_fr_to_en),
                },
            )
            self.wait_window(dlg)
            if dlg.result:
                d["dictionary_identification"] = dlg.result

        if "pure-transformationRules-regulatorCode.xlsx" in fmap:
            dlg = IdentificationDialog(
                self, "TransformationRules – Identification",
                [
                    ("version", True, tr_id.get("version", "")),
                    ("validityStartingDate", True, tr_id.get("validityStartingDate", "")),
                    ("description (EN)", True, tr_id.get("description (EN)", "")),
                    ("description (FR)", False, tr_id.get("description (FR)", "")),
                ],
                auto_translate={
                    "description (FR)": ("description (EN)", _translate_en_to_fr),
                },
            )
            self.wait_window(dlg)
            if dlg.result:
                d["transformation_rules_identification"] = dlg.result

        if "pure-transformedDataStructure-regulatoryDomainCode.xlsm" in fmap:
            dlg = IdentificationDialog(
                self, "TransformedDataStructure – Identification",
                [
                    ("version", True, td_id.get("version", "")),
                    ("validityStartingDate", True, td_id.get("validityStartingDate", "")),
                    ("description (EN)", True, td_id.get("description (EN)", "")),
                    ("description (FR)", False, td_id.get("description (FR)", "")),
                ],
                auto_translate={
                    "description (FR)": ("description (EN)", _translate_en_to_fr),
                },
            )
            self.wait_window(dlg)
            if dlg.result:
                d["transformed_data_identification"] = dlg.result

        if "pure-uiDashboard-regulatoryDomainCode.xlsx" in fmap:
            dlg = IdentificationDialog(
                self, "UiDashboard – Identification",
                [
                    ("version", True, ui_id.get("version", "")),
                    ("validityStartingDate", True, ui_id.get("validityStartingDate", "")),
                    ("description (EN)", True, ui_id.get("description (EN)", "")),
                    ("description (FR)", False, ui_id.get("description (FR)", "")),
                ],
                auto_translate={
                    "description (FR)": ("description (EN)", _translate_en_to_fr),
                },
            )
            self.wait_window(dlg)
            if dlg.result:
                d["ui_dashboard_identification"] = dlg.result

    def _open_output_identification_dialogs(self):
        d = self.collected
        fmap = d["output_file_map"]
        out_id = d.get("output_identification", {})
        out_struct_id = d.get("output_structure_identification", {})

        if "pure-regulatoryOutput-regulationCode-regulatoryOutputCode.xlsx" in fmap:
            dlg = IdentificationDialog(
                self, "Output – Identification + SubmissionFileset",
                [
                    ("code", True, out_id.get("code", "")),
                    ("version", True, out_id.get("version", "")),
                    ("validityStartingDate", True, out_id.get("validityStartingDate", "")),
                    ("periodicity", True, out_id.get("periodicity", "")),
                    ("commonCollection", True, out_id.get("commonCollection", "")),
                    ("name (EN)", True, out_id.get("name (EN)", "")),
                    ("name (FR)", False, out_id.get("name (FR)", "")),
                    ("description (EN)", True, out_id.get("description (EN)", "")),
                    ("description (FR)", False, out_id.get("description (FR)", "")),
                    ("submissionFormat", True, out_id.get("submissionFormat", "")),
                ],
                auto_translate={
                    "name (FR)": ("name (EN)", _translate_en_to_fr),
                    "description (FR)": ("description (EN)", _translate_en_to_fr),
                },
            )
            self.wait_window(dlg)
            if dlg.result:
                d["output_identification"] = dlg.result

        if "pure-outputStructure-regulatoryOutputCode-outputStructureCode.xlsx" in fmap:
            dlg = IdentificationDialog(
                self, "OutputStructure – Identification",
                [
                    ("version", True, out_struct_id.get("version", "")),
                    ("validityStartingDate", True, out_struct_id.get("validityStartingDate", "")),
                    ("name (EN)", True, out_struct_id.get("name (EN)", "")),
                    ("name (FR)", False, out_struct_id.get("name (FR)", "")),
                    ("description (EN)", True, out_struct_id.get("description (EN)", "")),
                    ("description (FR)", False, out_struct_id.get("description (FR)", "")),
                ],
                auto_translate={
                    "name (FR)": ("name (EN)", _translate_en_to_fr),
                    "description (FR)": ("description (EN)", _translate_en_to_fr),
                },
            )
            self.wait_window(dlg)
            if dlg.result:
                d["output_structure_identification"] = dlg.result


class LauncherApp(tk.Tk):
    """Initial launcher window: choose between Jira PDF import or manual mode."""

    def __init__(self):
        super().__init__()
        self.title("Regulation Project Initializer")
        self.geometry("560x480")
        self.configure(bg=BG)
        self.resizable(False, False)
        self.choice = None  # "pdf" or "manual"

        # Center window
        self.update_idletasks()
        x = (self.winfo_screenwidth() - 560) // 2
        y = (self.winfo_screenheight() - 480) // 2
        self.geometry(f"560x480+{x}+{y}")

        # Top accent bar
        tk.Frame(self, bg=ACCENT, height=4).pack(fill="x")

        # Title section
        title_section = tk.Frame(self, bg=BG)
        title_section.pack(fill="x", pady=(28, 0))
        tk.Label(
            title_section, text="⚙",
            font=("Segoe UI", 28), bg=BG, fg=ACCENT,
        ).pack()
        tk.Label(
            title_section, text="Regulation Project Initializer",
            font=("Segoe UI", 18, "bold"), bg=BG, fg=FG,
        ).pack(pady=(4, 2))
        tk.Label(
            title_section, text="How would you like to create the project?",
            font=("Segoe UI", 10), bg=BG, fg=SUBTLE_FG,
        ).pack(pady=(2, 20))

        # Cards container
        cards_frame = tk.Frame(self, bg=BG)
        cards_frame.pack(fill="x", padx=36)

        # Card 1: Jira PDF
        card1_border = tk.Frame(cards_frame, bg=ACCENT, padx=2, pady=2)
        card1_border.pack(fill="x", pady=(0, 14))
        card1 = tk.Frame(card1_border, bg=CARD_BG, cursor="hand2")
        card1.pack(fill="x")

        card1_content = tk.Frame(card1, bg=CARD_BG, padx=20, pady=16)
        card1_content.pack(fill="x")
        tk.Label(
            card1_content, text="📄  Import a Jira ticket (PDF)",
            font=("Segoe UI", 12, "bold"), bg=CARD_BG, fg=FG, anchor="w",
        ).pack(anchor="w")
        tk.Label(
            card1_content, text="Automatic generation from exported Jira tickets",
            font=("Segoe UI", 9), bg=CARD_BG, fg=SUBTLE_FG, anchor="w",
        ).pack(anchor="w", pady=(4, 8))
        btn_pdf = tk.Button(
            card1_content, text="Start  →",
            font=("Segoe UI", 10, "bold"),
            bg=ACCENT, fg=BTN_FG,
            activebackground=ACCENT_HOVER, activeforeground=BTN_FG,
            relief="flat", cursor="hand2", padx=20, pady=5,
            command=self._choose_pdf,
        )
        btn_pdf.pack(anchor="w")
        _bind_hover(btn_pdf, ACCENT, ACCENT_HOVER)

        # Make entire card clickable
        for w in (card1, card1_content, *card1_content.winfo_children()):
            if not isinstance(w, tk.Button):
                w.bind("<Button-1>", lambda e: self._choose_pdf())

        # Card 2: Manual
        card2_border = tk.Frame(cards_frame, bg=BORDER_COLOR, padx=2, pady=2)
        card2_border.pack(fill="x", pady=(0, 14))
        card2 = tk.Frame(card2_border, bg=CARD_BG, cursor="hand2")
        card2.pack(fill="x")

        card2_content = tk.Frame(card2, bg=CARD_BG, padx=20, pady=16)
        card2_content.pack(fill="x")
        tk.Label(
            card2_content, text="✏️  Manual entry",
            font=("Segoe UI", 12, "bold"), bg=CARD_BG, fg=FG, anchor="w",
        ).pack(anchor="w")
        tk.Label(
            card2_content, text="Step-by-step interactive wizard interface",
            font=("Segoe UI", 9), bg=CARD_BG, fg=SUBTLE_FG, anchor="w",
        ).pack(anchor="w", pady=(4, 8))
        btn_manual = tk.Button(
            card2_content, text="Start  →",
            font=("Segoe UI", 10, "bold"),
            bg=ENTRY_BG, fg=FG,
            activebackground=BORDER_COLOR, activeforeground=FG,
            relief="flat", cursor="hand2", padx=20, pady=5,
            command=self._choose_manual,
        )
        btn_manual.pack(anchor="w")
        _bind_hover(btn_manual, ENTRY_BG, BORDER_COLOR, FG, FG)

        for w in (card2, card2_content, *card2_content.winfo_children()):
            if not isinstance(w, tk.Button):
                w.bind("<Button-1>", lambda e: self._choose_manual())

        # Footer
        tk.Label(
            self, text="PURE Regulatory Compliance",
            font=("Segoe UI", 8), bg=BG, fg=SUBTLE_FG,
        ).pack(side="bottom", pady=(0, 12))

    def _choose_pdf(self):
        self.choice = "pdf"
        self.destroy()

    def _choose_manual(self):
        self.choice = "manual"
        self.destroy()


class JiraAutoApp(tk.Tk):
    """GUI for the automatic Jira PDF import mode."""

    def __init__(self):
        super().__init__()
        self.title("Regulation Project Initializer — Automatic Mode")
        self.geometry("880x750")
        self.configure(bg=BG)
        self.resizable(True, True)
        self.minsize(750, 600)

        self._jira_data = None
        self._collected = None
        self._result_dir = None

        # Top accent bar
        tk.Frame(self, bg=ACCENT, height=3).pack(fill="x")

        # Main container
        container = tk.Frame(self, bg=BG)
        container.pack(fill="both", expand=True, padx=22, pady=12)

        # Title section
        title_frame = tk.Frame(container, bg=BG)
        title_frame.pack(fill="x", pady=(0, 6))
        tk.Label(
            title_frame, text="⚙  Automatic Mode — Jira Ticket",
            font=("Segoe UI", 15, "bold"), bg=BG, fg=HEADER_FG,
        ).pack(side="left")
        # Status badge — soft neutral style
        self._status_badge = tk.Label(
            title_frame, text="  Waiting for file  ",
            font=("Segoe UI", 8, "bold"), bg=BORDER_COLOR, fg=FG,
            padx=8, pady=2,
        )
        self._status_badge.pack(side="right")

        ttk.Separator(container).pack(fill="x", pady=(0, 10))

        # File selection card
        file_card_border = tk.Frame(container, bg=BORDER_COLOR, padx=1, pady=1)
        file_card_border.pack(fill="x", pady=(0, 12))
        file_card = tk.Frame(file_card_border, bg=CARD_BG, padx=16, pady=12)
        file_card.pack(fill="x")

        btn_frame = tk.Frame(file_card, bg=CARD_BG)
        btn_frame.pack(fill="x")
        self._btn_select = tk.Button(
            btn_frame, text="📄  Select the Jira PDF file",
            font=("Segoe UI", 11, "bold"),
            bg=ACCENT, fg=BTN_FG, activebackground=ACCENT_HOVER,
            relief="flat", cursor="hand2", padx=16, pady=7,
            command=self._select_pdf,
        )
        self._btn_select.pack(side="left")
        _bind_hover(self._btn_select, ACCENT, ACCENT_HOVER)

        self._lbl_file = tk.Label(
            btn_frame, text="No file selected",
            font=("Segoe UI", 10, "italic"), bg=CARD_BG, fg=SUBTLE_FG,
        )
        self._lbl_file.pack(side="left", padx=14)

        # Destination folder chooser
        dest_header = tk.Frame(container, bg=BG)
        dest_header.pack(fill="x", pady=(0, 4))
        tk.Frame(dest_header, bg=ACCENT, width=4, height=16).pack(side="left", padx=(0, 8))
        tk.Label(
            dest_header, text="Destination Folder", font=("Segoe UI", 10, "bold"),
            bg=BG, fg=ACCENT,
        ).pack(side="left")

        dest_card_border = tk.Frame(container, bg=BORDER_COLOR, padx=1, pady=1)
        dest_card_border.pack(fill="x", pady=(0, 12))
        dest_card = tk.Frame(dest_card_border, bg=CARD_BG, padx=16, pady=10)
        dest_card.pack(fill="x")

        dest_row = tk.Frame(dest_card, bg=CARD_BG)
        dest_row.pack(fill="x")
        self._var_output_dir_auto = tk.StringVar(value=DEFAULT_OUTPUT_DIR)
        border_dest, _ = _make_styled_entry(dest_row, self._var_output_dir_auto, width=50)
        border_dest.pack(side="left", fill="x", expand=True)

        btn_browse_auto = tk.Button(
            dest_row, text="\U0001f4c2 Browse", font=("Segoe UI", 9),
            bg=BTN_BG, fg=BTN_FG, activebackground=ACCENT_HOVER, activeforeground=BTN_FG,
            relief="flat", padx=10, pady=4, cursor="hand2",
            command=self._browse_output_dir_auto,
        )
        btn_browse_auto.pack(side="left", padx=(6, 0))
        _bind_hover(btn_browse_auto, BTN_BG, ACCENT_HOVER)

        # Summary text area
        summary_header = tk.Frame(container, bg=BG)
        summary_header.pack(fill="x", pady=(0, 4))
        tk.Frame(summary_header, bg=ACCENT, width=4, height=16).pack(side="left", padx=(0, 8))
        tk.Label(
            summary_header, text="Ticket Summary", font=("Segoe UI", 10, "bold"),
            bg=BG, fg=ACCENT,
        ).pack(side="left")
        summary_border = tk.Frame(container, bg=BORDER_COLOR, padx=1, pady=1)
        summary_border.pack(fill="both", expand=True, pady=(0, 10))
        self._summary = scrolledtext.ScrolledText(
            summary_border, font=("Consolas", 10),
            bg=ENTRY_BG, fg=ENTRY_FG,
            insertbackground=FG, relief="flat",
            height=16, bd=0,
        )
        self._summary.pack(fill="both", expand=True)
        self._summary.config(state="disabled")

        # Action buttons frame
        ttk.Separator(container).pack(fill="x", pady=(0, 8))
        action_frame = tk.Frame(container, bg=BG)
        action_frame.pack(fill="x", pady=(0, 10))

        self._btn_generate = tk.Button(
            action_frame, text="🚀  Generate files",
            font=("Segoe UI", 11, "bold"),
            bg=BTN_BG, fg=BTN_FG,
            activebackground=ACCENT_HOVER,
            relief="flat", cursor="hand2", padx=16, pady=7,
            state="disabled",
            command=self._generate,
        )
        self._btn_generate.pack(side="left", padx=(0, 10))

        self._btn_push = tk.Button(
            action_frame, text="☁️  Push to GitLab",
            font=("Segoe UI", 11),
            bg=ENTRY_BG, fg=FG,
            activebackground=BORDER_COLOR,
            relief="flat", cursor="hand2", padx=16, pady=7,
            state="disabled",
            command=self._push_gitlab,
        )
        self._btn_push.pack(side="left")
        _bind_hover(self._btn_push, ENTRY_BG, BORDER_COLOR, FG, FG)

        self._btn_quit = tk.Button(
            action_frame, text="✖  Quit",
            font=("Segoe UI", 10),
            bg=BORDER_COLOR, fg=FG,
            activebackground=STEP_INACTIVE,
            relief="flat", cursor="hand2", padx=14, pady=7,
            command=self.destroy,
        )
        self._btn_quit.pack(side="right")
        _bind_hover(self._btn_quit, BORDER_COLOR, STEP_INACTIVE, FG, FG)

        # Log area
        log_header = tk.Frame(container, bg=BG)
        log_header.pack(fill="x", pady=(4, 4))
        tk.Frame(log_header, bg=SUBTLE_FG, width=4, height=16).pack(side="left", padx=(0, 8))
        tk.Label(
            log_header, text="Log", font=("Segoe UI", 10, "bold"),
            bg=BG, fg=SUBTLE_FG,
        ).pack(side="left")
        log_border = tk.Frame(container, bg=BORDER_COLOR, padx=1, pady=1)
        log_border.pack(fill="both", expand=True)
        self._log = scrolledtext.ScrolledText(
            log_border, font=("Consolas", 9),
            bg=ENTRY_BG, fg=SUBTLE_FG,
            insertbackground=FG, relief="flat",
            height=7, bd=0,
        )
        self._log.pack(fill="both", expand=True)
        self._log.config(state="disabled")

    def _log_msg(self, msg: str, tag: str = ""):
        self._log.config(state="normal")
        self._log.insert("end", msg + "\n", tag)
        self._log.see("end")
        self._log.config(state="disabled")

    def _browse_output_dir_auto(self):
        """Open a folder chooser for the project destination."""
        current = self._var_output_dir_auto.get().strip()
        initial = current if current and Path(current).exists() else str(Path.home())
        chosen = filedialog.askdirectory(title="Select destination folder", initialdir=initial)
        if chosen:
            self._var_output_dir_auto.set(chosen)

    def _select_pdf(self):
        path = filedialog.askopenfilename(
            parent=self,
            title="Select the Jira ticket (PDF)",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        )
        if not path:
            return

        self._lbl_file.config(text=Path(path).name, fg=ACCENT, font=("Segoe UI", 10))
        self._log_msg(f"File selected: {path}")
        self._status_badge.config(text="  Parsing...  ", bg=ACCENT, fg=BTN_FG)

        try:
            self._jira_data = parse_jira_file(path)
            self._collected = jira_data_to_collected(self._jira_data)

            summary = format_jira_summary(self._jira_data)
            self._summary.config(state="normal")
            self._summary.delete("1.0", "end")
            self._summary.insert("1.0", summary)
            self._summary.config(state="disabled")

            self._btn_generate.config(state="normal")
            _bind_hover(self._btn_generate, BTN_BG, ACCENT_HOVER)
            self._log_msg("Extraction completed successfully.")
            self._status_badge.config(text="  Ready to generate  ", bg=ACCENT, fg=BTN_FG)

        except Exception as exc:
            messagebox.showerror("Error", f"Unable to read the PDF:\n{exc}")
            self._log_msg(f"ERROR: {exc}")
            self._status_badge.config(text="  Error  ", bg=BORDER_COLOR, fg=FG)

    def _generate(self):
        if not self._collected:
            return

        self._btn_generate.config(state="disabled")
        self._log_msg("Generation in progress...")
        self.update_idletasks()

        try:
            collected = self._collected
            data = self._jira_data
            code = collected.get("regulation_code", "")
            project_type = collected.get("project_type", "both")
            git_name = _sanitize_project_name(collected.get("git_project_name", "") or code.lower())
            jira_mappings = collected.get("file_mappings", {})

            # Build file maps
            reg_file_map = {}
            out_file_map = {}

            if project_type in ("both", "regulation"):
                ref_dir = _reference_dir_for_scope("regulation")
                if ref_dir.exists():
                    for f in sorted(ref_dir.iterdir()):
                        if f.is_file():
                            target = _resolve_jira_file_name(f.name, code, jira_mappings)
                            final = _sanitize_file_name(target, f.suffix)
                            if final:
                                reg_file_map[f.name] = final

            if project_type in ("both", "output"):
                ref_dir = _reference_dir_for_scope("output")
                if ref_dir.exists():
                    for f in sorted(ref_dir.iterdir()):
                        if f.is_file():
                            target = _resolve_jira_file_name(f.name, code, jira_mappings)
                            final = _sanitize_file_name(target, f.suffix)
                            if final:
                                out_file_map[f.name] = final

            output_dir = self._var_output_dir_auto.get().strip() or DEFAULT_OUTPUT_DIR
            config = AppConfig(
                output_base_dir=Path(output_dir),
                project_base_name=git_name,
                project_type=project_type,
                regulation_folder_name="",
                output_folder_name="",
                regulation_file_map=reg_file_map,
                output_file_map=out_file_map,
                regulation_identification=collected.get("regulation_identification", {}),
                dictionary_identification=collected.get("dictionary_identification", {}),
                output_identification=collected.get("output_identification", {}),
                transformation_rules_identification=collected.get("transformation_rules_identification", {}),
                transformed_data_identification=collected.get("transformed_data_identification", {}),
                ui_dashboard_identification=collected.get("ui_dashboard_identification", {}),
                output_structure_identification=collected.get("output_structure_identification", {}),
                regulation_project_name=collected.get("regulation_project_name", ""),
                output_project_name=collected.get("output_project_name", ""),
            )

            regulation = RegulationInfo(
                code=code, version=data.version, name="",
                authority=data.authority_name, frequency="",
                first_reporting_date="",
            )
            project_input = ProjectInput(regulation=regulation)

            initializer = ProjectInitializer(config, project_input)
            self._result_dir = initializer.run()

            self._log_msg(f"Files generated in: {self._result_dir.resolve()}")

            # List generated files
            for p in sorted(self._result_dir.rglob("*")):
                if p.is_file():
                    rel = p.relative_to(self._result_dir)
                    self._log_msg(f"  ✓ {rel}")

            self._btn_push.config(state="normal")
            messagebox.showinfo("Success", f"Files generated in:\n{self._result_dir.resolve()}")

        except Exception as exc:
            self._log_msg(f"ERROR: {exc}")
            messagebox.showerror("Error", f"Error during generation:\n{exc}")
            self._btn_generate.config(state="normal")

    def _push_gitlab(self):
        if not self._result_dir:
            return

        # Ask for branch and token in a dialog
        dlg = GitLabPushDialog(self, self._collected, self._result_dir)
        self.wait_window(dlg)
        if dlg.success:
            self._log_msg("GitLab push completed successfully.")
        elif dlg.error:
            self._log_msg(f"GitLab error: {dlg.error}")


class GitLabPushDialog(tk.Toplevel):
    """Dialog for GitLab push settings."""

    def __init__(self, parent, collected, result_dir):
        super().__init__(parent)
        self.title("Push to GitLab")
        self.geometry("540x400")
        self.configure(bg=BG)
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        self.success = False
        self.error = None
        self._collected = collected
        self._result_dir = result_dir

        # Center on parent
        self.update_idletasks()
        px = parent.winfo_rootx() + (parent.winfo_width() - 540) // 2
        py = parent.winfo_rooty() + (parent.winfo_height() - 400) // 2
        self.geometry(f"540x400+{px}+{py}")

        # Top accent bar
        tk.Frame(self, bg=ACCENT, height=3).pack(fill="x")

        # Title
        title_frame = tk.Frame(self, bg=BG)
        title_frame.pack(fill="x", padx=20, pady=(16, 4))
        tk.Frame(title_frame, bg=ACCENT, width=4, height=20).pack(side="left", padx=(0, 10))
        tk.Label(title_frame, text="GitLab Configuration", font=("Segoe UI", 13, "bold"),
                 bg=BG, fg=HEADER_FG).pack(side="left")

        ttk.Separator(self).pack(fill="x", padx=20, pady=(8, 12))

        form = tk.Frame(self, bg=BG, padx=20)
        form.pack(fill="both", expand=True)

        # Branch
        tk.Label(form, text="Target branch:", font=("Segoe UI", 10, "bold"),
                 bg=BG, fg=FG).pack(anchor="w", pady=(0, 4))
        self._var_branch = tk.StringVar(value="main")
        b1, _ = _make_styled_entry(form, self._var_branch, font=("Segoe UI", 10))
        b1.pack(fill="x", pady=(0, 12))

        # Token
        tk.Label(form, text="GitLab Token (private):", font=("Segoe UI", 10, "bold"),
                 bg=BG, fg=FG).pack(anchor="w", pady=(0, 4))
        self._var_token = tk.StringVar(value=os.environ.get("GITLAB_TOKEN", ""))
        b2, _ = _make_styled_entry(form, self._var_token, show="●", font=("Segoe UI", 10))
        b2.pack(fill="x", pady=(0, 12))

        # Project name
        base_name = _sanitize_project_name(collected.get("git_project_name", ""))
        project_type = collected.get("project_type", "both")
        git_namespace = _extract_gitlab_namespace(collected.get("git_project_url", ""))
        reg_proj_name = collected.get("regulation_project_name") or f"pure-regulation-{base_name}"
        out_proj_name = collected.get("output_project_name") or f"pure-regulatoryoutput-{base_name}"
        # Prefix with namespace if available, so user sees the full target path
        if git_namespace:
            reg_proj_name = f"{git_namespace}/{reg_proj_name}"
            out_proj_name = f"{git_namespace}/{out_proj_name}"

        if project_type in ("both", "regulation"):
            tk.Label(form, text="Regulation project name:", font=("Segoe UI", 10, "bold"),
                     bg=BG, fg=FG).pack(anchor="w", pady=(0, 4))
            self._var_reg_name = tk.StringVar(value=reg_proj_name)
            b3, _ = _make_styled_entry(form, self._var_reg_name, font=("Segoe UI", 10))
            b3.pack(fill="x", pady=(0, 12))
        else:
            self._var_reg_name = None

        if project_type in ("both", "output"):
            tk.Label(form, text="Output project name:", font=("Segoe UI", 10, "bold"),
                     bg=BG, fg=FG).pack(anchor="w", pady=(0, 4))
            self._var_out_name = tk.StringVar(value=out_proj_name)
            b4, _ = _make_styled_entry(form, self._var_out_name, font=("Segoe UI", 10))
            b4.pack(fill="x", pady=(0, 12))
        else:
            self._var_out_name = None

        # Push button
        ttk.Separator(self).pack(fill="x", padx=20)
        btn_frame = tk.Frame(self, bg=BG)
        btn_frame.pack(fill="x", padx=20, pady=14)
        btn_push = tk.Button(
            btn_frame, text="🚀  Push", font=("Segoe UI", 11, "bold"),
            bg=ACCENT, fg=BTN_FG, activebackground=ACCENT_HOVER,
            relief="flat", cursor="hand2", padx=20, pady=7,
            command=self._push,
        )
        btn_push.pack(side="right")
        _bind_hover(btn_push, ACCENT, ACCENT_HOVER)

    def _push(self):
        token = self._var_token.get().strip()
        branch = self._var_branch.get().strip() or "main"

        if not token:
            messagebox.showwarning("Token required", "Enter your GitLab token.", parent=self)
            return

        try:
            gitlab_cfg = GitLabConfig(private_token=token)
            publisher = GitLabPublisher(gitlab_cfg)
            gl_user = publisher.get_authenticated_username()

            # Use the namespace from Jira git_project_url if available
            git_namespace = _extract_gitlab_namespace(self._collected.get("git_project_url", ""))

            project_type = self._collected.get("project_type", "both")
            reg = self._collected.get("regulation_code", "")
            base = _sanitize_project_name(self._collected.get("git_project_name", ""))
            reg_folder = self._collected.get("regulation_project_name") or f"pure-regulation-{base}"
            out_folder = self._collected.get("output_project_name") or f"pure-regulatoryoutput-{base}"

            if project_type in ("both", "regulation") and self._var_reg_name:
                reg_name = self._var_reg_name.get().strip()
                if "/" in reg_name:
                    reg_project = reg_name
                elif git_namespace:
                    reg_project = f"{git_namespace}/{reg_name}"
                else:
                    reg_project = f"{gl_user}/{reg_name}"
                publisher.ensure_project_exists(reg_project, gl_user, create_if_missing=True)
                reg_dir = self._result_dir / reg_folder
                if reg_dir.exists():
                    publisher.publish_directory(reg_project, reg_dir, branch,
                                                f"chore: init {reg} regulation")

            if project_type in ("both", "output") and self._var_out_name:
                out_name = self._var_out_name.get().strip()
                if "/" in out_name:
                    out_project = out_name
                elif git_namespace:
                    out_project = f"{git_namespace}/{out_name}"
                else:
                    out_project = f"{gl_user}/{out_name}"
                publisher.ensure_project_exists(out_project, gl_user, create_if_missing=True)
                out_dir = self._result_dir / out_folder
                if out_dir.exists():
                    publisher.publish_directory(out_project, out_dir, branch,
                                                f"chore: init {reg} output")

            self.success = True
            messagebox.showinfo("Success", "Push to GitLab completed.", parent=self)
            self.destroy()

        except Exception as exc:
            self.error = str(exc)
            messagebox.showerror("GitLab Error", str(exc), parent=self)


def _resolve_jira_file_name(file_name: str, regulation_code: str, jira_mappings: dict) -> str:
    """Resolve a filename using Jira mappings or default tokenization."""
    # Try exact match first
    if file_name in jira_mappings:
        return jira_mappings[file_name]
    # Try case-insensitive exact match
    for jira_src, jira_dst in jira_mappings.items():
        if jira_src.lower() == file_name.lower():
            return jira_dst
    # Try prefix match (everything before the last '-')
    for jira_src, jira_dst in jira_mappings.items():
        jira_prefix = jira_src.rsplit("-", 1)[0] if "-" in jira_src else jira_src
        ref_prefix = file_name.rsplit("-", 1)[0] if "-" in file_name else file_name
        if jira_prefix.lower() == ref_prefix.lower():
            return jira_dst
    return _tokenized_default_name(file_name, regulation_code)


def run_gui():
    """Entry point for the GUI."""
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s – %(message)s",
        datefmt="%H:%M:%S",
    )

    # Show launcher to choose mode
    launcher = LauncherApp()
    launcher.mainloop()

    if launcher.choice == "pdf":
        app = JiraAutoApp()
        app.mainloop()
    elif launcher.choice == "manual":
        app = RegulationInitAppFull(manual_only=True)
        app.mainloop()


if __name__ == "__main__":
    run_gui()
