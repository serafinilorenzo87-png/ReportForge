from __future__ import annotations

from collections import defaultdict, OrderedDict
from datetime import datetime
from pathlib import Path
from tempfile import TemporaryDirectory
import shutil
import webbrowser

from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QStackedWidget,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from app.models.checklist_item import ChecklistItem
from app.models.finding import Finding
from app.models.project import Project
from app.models.reference import Reference
from app.services.database import SessionLocal
from app.services.evidence_manager import copy_evidence_files
from app.services.finding_templates import get_template_data, get_template_names
from app.services.reference_helpers import normalize_url, smart_fill_reference_fields
from app.services.report_exporter import (
    export_docx_report,
    export_markdown_report,
    export_pdf_report,
    export_text_report,
)
from app.ui.styles import APP_STYLE


CHECKLIST_SECTIONS: list[tuple[str, list[tuple[str, str]]]] = [
    (
        "Pre-Engagement",
        [
            ("scope_confirmed", "Scope confirmed"),
            ("targets_confirmed", "Targets confirmed"),
            ("out_of_scope_documented", "Out-of-scope systems documented"),
            ("test_window_confirmed", "Test window confirmed"),
            ("rules_of_engagement_defined", "Rules of engagement defined"),
            ("stakeholder_contacts_available", "Stakeholder contacts available"),
            ("required_credentials_received", "Required credentials received"),
            ("vpn_access_tested", "VPN / access path tested"),
            ("evidence_storage_ready", "Evidence storage path ready"),
        ],
    ),
    (
        "Recon & Enumeration",
        [
            ("dns_enumeration_completed", "DNS / domain enumeration completed"),
            ("subdomain_discovery_completed", "Subdomain discovery completed"),
            ("port_scanning_completed", "Port scanning completed"),
            ("service_enumeration_completed", "Service enumeration completed"),
            ("technology_fingerprinting_completed", "Technology fingerprinting completed"),
            ("web_content_discovery_completed", "Web content discovery completed"),
            ("authentication_surface_mapped", "Authentication surface mapped"),
            ("attack_surface_documented", "Attack surface documented"),
        ],
    ),
    (
        "Validation & Exploitation",
        [
            ("authentication_weaknesses_tested", "Authentication weaknesses tested"),
            ("authorization_weaknesses_tested", "Authorization weaknesses tested"),
            ("input_validation_tested", "Input validation tested"),
            ("file_upload_tested", "File upload tested"),
            ("session_management_tested", "Session management tested"),
            ("sensitive_data_exposure_checked", "Sensitive data exposure checked"),
            ("misconfigurations_validated", "Misconfigurations validated"),
            ("exploitation_documented", "Exploitation attempts documented"),
            ("lateral_movement_reviewed", "Lateral movement potential reviewed"),
        ],
    ),
    (
        "Post-Exploitation & Impact",
        [
            ("privilege_level_documented", "Privilege level documented"),
            ("data_exposure_impact_reviewed", "Data exposure impact reviewed"),
            ("persistence_risk_considered", "Persistence risk considered"),
            ("business_impact_considered", "Business impact considered"),
            ("cleanup_completed", "Cleanup completed"),
            ("artifacts_removed", "Artifacts removed where required"),
        ],
    ),
    (
        "Reporting Quality Check",
        [
            ("all_findings_have_severity", "All findings have severity"),
            ("all_findings_have_description", "All findings have description"),
            ("all_findings_have_evidence", "All findings have evidence"),
            ("all_findings_have_remediation", "All findings have remediation"),
            ("references_added", "References added where relevant"),
            ("screenshots_attached", "Screenshots attached where relevant"),
            ("stakeholder_summary_completed", "Stakeholder summary completed"),
            ("executive_summary_completed", "Executive summary completed"),
            ("technical_summary_completed", "Technical summary completed"),
            ("conclusions_completed", "Conclusions completed"),
            ("risk_rating_reviewed", "Risk rating reviewed"),
        ],
    ),
    (
        "Final Delivery Readiness",
        [
            ("project_metadata_reviewed", "Project metadata reviewed"),
            ("dates_reviewed", "Dates reviewed"),
            ("ips_domains_reviewed", "IPs / domains reviewed"),
            ("application_names_reviewed", "Application names reviewed"),
            ("methodology_reviewed", "Methodology section reviewed"),
            ("standards_reviewed", "Standards section reviewed"),
            ("severity_chart_verified", "Severity chart verified"),
            ("docx_export_checked", "DOCX export checked"),
            ("pdf_export_checked", "PDF export checked"),
            ("final_report_ready", "Final report ready"),
        ],
    ),
]


class ClickableProjectCard(QFrame):
    def __init__(self, project_id: int, on_click_callback):
        super().__init__()
        self.project_id = project_id
        self.on_click_callback = on_click_callback
        self.setCursor(Qt.CursorShape.PointingHandCursor)

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.on_click_callback(self.project_id)
        super().mousePressEvent(event)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("ReportForge")
        icon_path = Path(__file__).resolve().parents[2] / "assets" / "reportforge_icon.png"
        if icon_path.exists():
            self.setWindowIcon(QIcon(str(icon_path)))
        self.setMinimumSize(1500, 900)
        self.setStyleSheet(APP_STYLE)

        self.current_project_id: int | None = None
        self.current_edit_project_id: int | None = None
        self.current_edit_finding_id: int | None = None
        self.current_edit_reference_id: int | None = None
        self.current_selected_finding_id: int | None = None
        self.current_finding_attachments: list[str] = []
        self.recent_projects_container_layout = None
        self.project_archive_groups = OrderedDict()
        self.selected_archive_filter = "Recent"
        self.all_projects_cache: list[Project] = []
        self.last_export_directory: Path | None = None

        self.checklist_combo_by_key: dict[str, QComboBox] = {}
        self.checklist_section_cards: dict[str, QFrame] = {}
        self.checklist_row_by_key: dict[str, QFrame] = {}
        self.checklist_label_by_key: dict[str, QLabel] = {}
        self.checklist_section_by_key: dict[str, str] = {}
        self.pending_navigation_active = False
        self.current_pending_key: str | None = None

        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        root_layout = QHBoxLayout()
        root_layout.setContentsMargins(16, 16, 16, 16)
        root_layout.setSpacing(16)

        sidebar = QFrame()
        sidebar.setObjectName("Card")
        sidebar.setFixedWidth(250)

        sidebar_layout = QVBoxLayout()
        sidebar_layout.setContentsMargins(18, 18, 18, 18)
        sidebar_layout.setSpacing(14)

        app_title = QLabel("ReportForge")
        app_title.setStyleSheet("font-size: 22px; font-weight: bold;")

        self.nav_list = QListWidget()
        self.nav_items = [
            "Dashboard",
            "Projects",
            "Findings",
            "References",
            "Methodology Checklist",
            "Export",
        ]

        for item_text in self.nav_items:
            QListWidgetItem(item_text, self.nav_list)

        self.nav_list.setCurrentRow(0)

        sidebar_layout.addWidget(app_title)
        sidebar_layout.addWidget(self.nav_list)
        sidebar.setLayout(sidebar_layout)

        self.pages = QStackedWidget()
        self.pages.addWidget(self.build_dashboard_page())
        self.pages.addWidget(self.build_projects_page())
        self.pages.addWidget(self.build_findings_page())
        self.pages.addWidget(self.build_references_page())
        self.pages.addWidget(self.build_checklist_page())
        self.pages.addWidget(self.build_export_page())

        self.nav_list.currentRowChanged.connect(self.pages.setCurrentIndex)

        root_layout.addWidget(sidebar)
        root_layout.addWidget(self.pages, 1)

        footer_label = QLabel("Created by Lorenzo Serafini")
        footer_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        footer_label.setStyleSheet(
            "font-size: 13px; font-weight: 600; color: #64748b; padding: 8px 0 2px 0;"
        )

        outer_layout = QVBoxLayout()
        outer_layout.setContentsMargins(0, 0, 0, 0)
        outer_layout.setSpacing(6)
        outer_layout.addLayout(root_layout, 1)
        outer_layout.addWidget(footer_label)

        central_widget.setLayout(outer_layout)

        self.load_projects()
        self.load_findings()
        self.load_reference_finding_list()
        self.load_references()
        self.update_checklist_context()
        self.update_export_context()

    def format_project_date(self, created_at: datetime | None) -> str:
        if created_at is None:
            return "Unknown date"
        return created_at.strftime("%d %B %Y")

    def format_project_month_group(self, created_at: datetime | None) -> str:
        if created_at is None:
            return "Unknown Month"
        return created_at.strftime("%B %Y")

    def severity_object_name(self, severity: str) -> str:
        normalized = (severity or "Info").strip().capitalize()
        mapping = {
            "Critical": "SeverityBadgeCritical",
            "High": "SeverityBadgeHigh",
            "Medium": "SeverityBadgeMedium",
            "Low": "SeverityBadgeLow",
            "Info": "SeverityBadgeInfo",
        }
        return mapping.get(normalized, "SeverityBadgeInfo")

    def create_input_label(self, text: str) -> QLabel:
        label = QLabel(text)
        label.setStyleSheet("font-size: 13px; font-weight: 600; color: #374151;")
        return label

    def create_multiline_input(self, placeholder: str, height: int = 90) -> QTextEdit:
        widget = QTextEdit()
        widget.setPlaceholderText(placeholder)
        widget.setFixedHeight(height)
        return widget

    def build_dashboard_page(self) -> QWidget:
        page = QWidget()

        content_layout = QVBoxLayout()
        content_layout.setContentsMargins(0, 0, 0, 0)
        content_layout.setSpacing(16)
        content_layout.setAlignment(Qt.AlignmentFlag.AlignTop)

        header_title = QLabel("DASHBOARD")
        header_title.setStyleSheet("font-size: 30px; font-weight: bold;")

        header_subtitle = QLabel(
            "Create, review, and organize polished local pentest reporting projects."
        )
        header_subtitle.setStyleSheet("font-size: 14px; color: #4b5563;")

        welcome_card = QFrame()
        welcome_card.setObjectName("Card")

        welcome_layout = QVBoxLayout()
        welcome_layout.setContentsMargins(24, 24, 24, 24)
        welcome_layout.setSpacing(12)

        welcome_title = QLabel("Welcome")
        welcome_title.setStyleSheet("font-size: 20px; font-weight: bold;")

        welcome_text = QLabel(
            "Use ReportForge to structure projects, document findings, keep references local, "
            "and export professional reports for both stakeholders and technical teams."
        )
        welcome_text.setWordWrap(True)
        welcome_text.setStyleSheet("font-size: 14px; color: #4b5563;")

        create_project_button = QPushButton("Create New Project")
        create_project_button.clicked.connect(self.go_to_new_project_form)

        welcome_layout.addWidget(welcome_title)
        welcome_layout.addWidget(welcome_text)
        welcome_layout.addWidget(create_project_button, alignment=Qt.AlignmentFlag.AlignLeft)
        welcome_card.setLayout(welcome_layout)

        recent_card = QFrame()
        recent_card.setObjectName("Card")

        recent_layout = QVBoxLayout()
        recent_layout.setContentsMargins(24, 24, 24, 24)
        recent_layout.setSpacing(12)

        recent_title = QLabel("Recent Projects")
        recent_title.setStyleSheet("font-size: 20px; font-weight: bold;")

        recent_subtitle = QLabel("Quick access to your latest project activity.")
        recent_subtitle.setStyleSheet("font-size: 13px; color: #6b7280;")

        recent_container = QWidget()
        self.recent_projects_container_layout = QVBoxLayout()
        self.recent_projects_container_layout.setContentsMargins(0, 0, 0, 0)
        self.recent_projects_container_layout.setSpacing(12)
        self.recent_projects_container_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        recent_container.setLayout(self.recent_projects_container_layout)

        recent_layout.addWidget(recent_title)
        recent_layout.addWidget(recent_subtitle)
        recent_layout.addWidget(recent_container)

        recent_card.setLayout(recent_layout)

        content_layout.addWidget(header_title)
        content_layout.addWidget(header_subtitle)
        content_layout.addWidget(welcome_card)
        content_layout.addWidget(recent_card)

        page.setLayout(content_layout)
        return page

    def build_projects_page(self) -> QWidget:
        page = QWidget()

        outer_layout = QVBoxLayout()
        outer_layout.setContentsMargins(0, 0, 0, 0)
        outer_layout.setSpacing(16)

        page_title = QLabel("Projects")
        page_title.setStyleSheet("font-size: 30px; font-weight: bold;")

        page_subtitle = QLabel(
            "Create a premium local pentest reporting project with engagement, scope, assets, summaries, and conclusions."
        )
        page_subtitle.setStyleSheet("font-size: 14px; color: #4b5563;")

        self.current_project_card = QFrame()
        self.current_project_card.setObjectName("Card")

        current_project_layout = QVBoxLayout()
        current_project_layout.setContentsMargins(24, 20, 24, 20)
        current_project_layout.setSpacing(6)

        current_project_title = QLabel("Current Project")
        current_project_title.setStyleSheet("font-size: 18px; font-weight: bold;")

        self.current_project_name_label = QLabel("No project selected")
        self.current_project_name_label.setStyleSheet(
            "font-size: 16px; font-weight: bold; color: #111827;"
        )

        self.current_project_date_label = QLabel("")
        self.current_project_date_label.setStyleSheet("font-size: 13px; color: #6b7280;")

        self.current_project_meta_label = QLabel(
            "Select a saved project to make it the active workspace."
        )
        self.current_project_meta_label.setWordWrap(True)
        self.current_project_meta_label.setStyleSheet("font-size: 13px; color: #6b7280;")

        current_project_layout.addWidget(current_project_title)
        current_project_layout.addWidget(self.current_project_name_label)
        current_project_layout.addWidget(self.current_project_date_label)
        current_project_layout.addWidget(self.current_project_meta_label)
        self.current_project_card.setLayout(current_project_layout)

        content_row = QHBoxLayout()
        content_row.setSpacing(20)
        content_row.setAlignment(Qt.AlignmentFlag.AlignTop)

        form_card = QFrame()
        form_card.setObjectName("Card")
        form_card.setMinimumWidth(760)
        form_card.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Maximum)

        form_layout = QVBoxLayout()
        form_layout.setContentsMargins(24, 24, 24, 24)
        form_layout.setSpacing(12)

        self.form_title = QLabel("New Project")
        self.form_title.setStyleSheet("font-size: 20px; font-weight: bold;")

        self.form_subtitle = QLabel(
            "Define the engagement, assets, summaries, and report content for a new assessment project."
        )
        self.form_subtitle.setWordWrap(True)
        self.form_subtitle.setStyleSheet("font-size: 13px; color: #6b7280;")

        grid = QGridLayout()
        grid.setHorizontalSpacing(14)
        grid.setVerticalSpacing(10)

        self.project_name_input = QLineEdit()
        self.project_name_input.setPlaceholderText("Example: External Pentest - Customer Portal")

        self.target_name_input = QLineEdit()
        self.target_name_input.setPlaceholderText("Example: customer.example.com")

        self.platform_input = QLineEdit()
        self.platform_input.setPlaceholderText("Example: Internal Infrastructure / Web App / HTB")

        self.client_name_input = QLineEdit()
        self.client_name_input.setPlaceholderText("Example: Acme Corp")

        self.engagement_type_input = QComboBox()
        self.engagement_type_input.addItems(
            [
                "External Pentest",
                "Internal Pentest",
                "Web Application Assessment",
                "Mobile Assessment",
                "API Assessment",
                "Red Team Style Exercise",
                "HTB / Lab Report",
                "Other",
            ]
        )

        self.assessment_start_input = QLineEdit()
        self.assessment_start_input.setInputMask("00/00/0000")
        self.assessment_start_input.setPlaceholderText("gg/mm/aaaa")

        self.assessment_end_input = QLineEdit()
        self.assessment_end_input.setInputMask("00/00/0000")
        self.assessment_end_input.setPlaceholderText("gg/mm/aaaa")

        self.risk_rating_input = QComboBox()
        self.risk_rating_input.addItems(["Not Rated", "Critical", "High", "Medium", "Low"])

        grid.addWidget(self.create_input_label("Project Name"), 0, 0)
        grid.addWidget(self.project_name_input, 1, 0)
        grid.addWidget(self.create_input_label("Target Name"), 0, 1)
        grid.addWidget(self.target_name_input, 1, 1)

        grid.addWidget(self.create_input_label("Client Name"), 2, 0)
        grid.addWidget(self.client_name_input, 3, 0)
        grid.addWidget(self.create_input_label("Platform"), 2, 1)
        grid.addWidget(self.platform_input, 3, 1)

        grid.addWidget(self.create_input_label("Engagement Type"), 4, 0)
        grid.addWidget(self.engagement_type_input, 5, 0)
        grid.addWidget(self.create_input_label("Risk Rating"), 4, 1)
        grid.addWidget(self.risk_rating_input, 5, 1)

        grid.addWidget(self.create_input_label("Assessment Start"), 6, 0)
        grid.addWidget(self.assessment_start_input, 7, 0)
        grid.addWidget(self.create_input_label("Assessment End"), 6, 1)
        grid.addWidget(self.assessment_end_input, 7, 1)

        self.target_ips_input = self.create_multiline_input("IPs, CIDRs, hosts, assets...", 70)
        self.target_domains_input = self.create_multiline_input("Domains, subdomains, hostnames...", 70)
        self.web_app_name_input = QLineEdit()
        self.web_app_name_input.setPlaceholderText("Example: Customer Portal")
        self.mobile_app_name_input = QLineEdit()
        self.mobile_app_name_input.setPlaceholderText("Example: iOS App")
        self.internal_app_name_input = QLineEdit()
        self.internal_app_name_input.setPlaceholderText("Example: Admin Panel")
        self.environment_name_input = QLineEdit()
        self.environment_name_input.setPlaceholderText("Example: Production")

        assets_grid = QGridLayout()
        assets_grid.setHorizontalSpacing(14)
        assets_grid.setVerticalSpacing(10)
        assets_grid.addWidget(self.create_input_label("Target IPs / Assets"), 0, 0)
        assets_grid.addWidget(self.target_ips_input, 1, 0)
        assets_grid.addWidget(self.create_input_label("Target Domains / Hostnames"), 0, 1)
        assets_grid.addWidget(self.target_domains_input, 1, 1)
        assets_grid.addWidget(self.create_input_label("Web Application"), 2, 0)
        assets_grid.addWidget(self.web_app_name_input, 3, 0)
        assets_grid.addWidget(self.create_input_label("Mobile Application"), 2, 1)
        assets_grid.addWidget(self.mobile_app_name_input, 3, 1)
        assets_grid.addWidget(self.create_input_label("Internal Application"), 4, 0)
        assets_grid.addWidget(self.internal_app_name_input, 5, 0)
        assets_grid.addWidget(self.create_input_label("Environment"), 4, 1)
        assets_grid.addWidget(self.environment_name_input, 5, 1)

        self.scope_summary_input = self.create_multiline_input("Describe scope, perimeter, objectives, and authorized targets.", 90)
        self.out_of_scope_input = self.create_multiline_input("Describe exclusions, forbidden systems, or limitations.", 80)
        self.attack_surface_input = self.create_multiline_input("Describe exposed services, entry points, internet-facing assets, apps, APIs, ports, identity surfaces, and trust boundaries.", 110)
        self.methodology_summary_input = self.create_multiline_input("Describe methodology, phases, tooling approach, and validation process.", 95)
        self.standards_used_input = self.create_multiline_input("Example: OWASP Testing Guide, PTES, NIST, CWE, CVE, internal methodology...", 80)
        self.stakeholder_summary_input = self.create_multiline_input("Business-facing high-level summary for non-technical stakeholders.", 100)
        self.executive_summary_input = self.create_multiline_input("Executive summary of overall assessment posture, risk, and recommended priorities.", 100)
        self.technical_summary_input = self.create_multiline_input("Technical summary for engineers and remediation owners.", 100)
        self.conclusions_input = self.create_multiline_input("Write strong final conclusions, next steps, and overall assessment posture.", 110)
        self.project_notes_input = self.create_multiline_input("Additional notes, assumptions, observations, or useful context.", 90)

        self.save_button = QPushButton("Save Project")
        self.save_button.clicked.connect(self.handle_save_project)
        self.save_button.setFixedWidth(140)

        self.clear_form_button = QPushButton("Clear Form")
        self.clear_form_button.clicked.connect(self.reset_project_form)
        self.clear_form_button.setObjectName("SecondaryButton")
        self.clear_form_button.setFixedWidth(120)

        self.update_button = QPushButton("Update Project")
        self.update_button.clicked.connect(self.handle_update_project)
        self.update_button.setFixedWidth(140)
        self.update_button.hide()

        self.cancel_edit_button = QPushButton("Cancel Editing")
        self.cancel_edit_button.clicked.connect(self.reset_project_form)
        self.cancel_edit_button.setObjectName("SecondaryButton")
        self.cancel_edit_button.setFixedWidth(130)
        self.cancel_edit_button.hide()

        button_row = QHBoxLayout()
        button_row.setSpacing(10)
        button_row.addWidget(self.save_button)
        button_row.addWidget(self.clear_form_button)
        button_row.addWidget(self.update_button)
        button_row.addWidget(self.cancel_edit_button)
        button_row.addStretch()

        form_layout.addWidget(self.form_title)
        form_layout.addWidget(self.form_subtitle)
        form_layout.addSpacing(4)
        form_layout.addLayout(grid)

        section_titles = [
            ("Targets, Applications, and Environment", assets_grid),
            ("Scope Summary", self.scope_summary_input),
            ("Out of Scope", self.out_of_scope_input),
            ("Attack Surface", self.attack_surface_input),
            ("Methodology", self.methodology_summary_input),
            ("Standards Used", self.standards_used_input),
            ("Stakeholder Summary", self.stakeholder_summary_input),
            ("Executive Summary", self.executive_summary_input),
            ("Technical Summary", self.technical_summary_input),
            ("Conclusions", self.conclusions_input),
            ("Additional Notes", self.project_notes_input),
        ]

        for title, widget in section_titles:
            section_title = QLabel(title)
            section_title.setStyleSheet("font-size: 15px; font-weight: bold; color: #111827; padding-top: 8px;")
            form_layout.addWidget(section_title)
            if isinstance(widget, QGridLayout):
                form_layout.addLayout(widget)
            else:
                form_layout.addWidget(widget)

        form_layout.addSpacing(8)
        form_layout.addLayout(button_row)
        form_card.setLayout(form_layout)

        form_scroll = QScrollArea()
        form_scroll.setWidgetResizable(True)
        form_scroll.setFrameShape(QFrame.Shape.NoFrame)
        form_scroll.setWidget(form_card)
        form_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)

        projects_card = QFrame()
        projects_card.setObjectName("Card")
        projects_card.setMinimumWidth(500)
        projects_card.setMaximumWidth(540)
        projects_card.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Expanding)

        projects_layout = QVBoxLayout()
        projects_layout.setContentsMargins(24, 24, 24, 24)
        projects_layout.setSpacing(12)

        projects_title = QLabel("Saved Projects")
        projects_title.setStyleSheet("font-size: 20px; font-weight: bold;")

        projects_subtitle = QLabel(
            "Browse your project history by recent activity or monthly archive."
        )
        projects_subtitle.setStyleSheet("font-size: 13px; color: #6b7280;")

        self.project_search_input = QLineEdit()
        self.project_search_input.setPlaceholderText(
            "Search projects by name, client, target, platform..."
        )
        self.project_search_input.textChanged.connect(self.render_projects_browser)

        browser_row = QHBoxLayout()
        browser_row.setSpacing(10)
        browser_row.setAlignment(Qt.AlignmentFlag.AlignTop)

        archive_card = QFrame()
        archive_card.setObjectName("ArchivePanel")
        archive_card.setFixedWidth(150)

        archive_layout = QVBoxLayout()
        archive_layout.setContentsMargins(14, 14, 14, 14)
        archive_layout.setSpacing(10)

        archive_title = QLabel("Archive")
        archive_title.setStyleSheet("font-size: 15px; font-weight: bold;")

        self.archive_list = QListWidget()
        self.archive_list.currentTextChanged.connect(self.handle_archive_filter_change)

        archive_layout.addWidget(archive_title)
        archive_layout.addWidget(self.archive_list)
        archive_card.setLayout(archive_layout)

        projects_scroll_area = QScrollArea()
        projects_scroll_area.setWidgetResizable(True)
        projects_scroll_area.setFrameShape(QFrame.Shape.NoFrame)

        self.projects_container = QWidget()
        self.projects_container_layout = QVBoxLayout()
        self.projects_container_layout.setContentsMargins(0, 0, 0, 0)
        self.projects_container_layout.setSpacing(14)
        self.projects_container_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        self.projects_container.setLayout(self.projects_container_layout)
        projects_scroll_area.setWidget(self.projects_container)

        browser_row.addWidget(archive_card)
        browser_row.addWidget(projects_scroll_area, 1)

        projects_layout.addWidget(projects_title)
        projects_layout.addWidget(projects_subtitle)
        projects_layout.addWidget(self.project_search_input)
        projects_layout.addLayout(browser_row)
        projects_card.setLayout(projects_layout)

        content_row.addWidget(form_scroll, 5)
        content_row.addWidget(projects_card, 2)
        content_row.setStretch(0, 5)
        content_row.setStretch(1, 2)
        projects_scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)

        outer_layout.addWidget(page_title)
        outer_layout.addWidget(page_subtitle)
        outer_layout.addWidget(self.current_project_card)
        outer_layout.addLayout(content_row)

        page.setLayout(outer_layout)
        return page

    def collect_project_form_data(self) -> dict:
        return {
            "project_name": self.project_name_input.text().strip(),
            "target_name": self.target_name_input.text().strip(),
            "platform": self.platform_input.text().strip(),
            "client_name": self.client_name_input.text().strip(),
            "engagement_type": self.engagement_type_input.currentText().strip(),
            "assessment_start": self.assessment_start_input.text().strip(),
            "assessment_end": self.assessment_end_input.text().strip(),
            "target_ips": self.target_ips_input.toPlainText().strip(),
            "target_domains": self.target_domains_input.toPlainText().strip(),
            "web_app_name": self.web_app_name_input.text().strip(),
            "mobile_app_name": self.mobile_app_name_input.text().strip(),
            "internal_app_name": self.internal_app_name_input.text().strip(),
            "environment_name": self.environment_name_input.text().strip(),
            "scope_summary": self.scope_summary_input.toPlainText().strip(),
            "out_of_scope": self.out_of_scope_input.toPlainText().strip(),
            "attack_surface": self.attack_surface_input.toPlainText().strip(),
            "methodology_summary": self.methodology_summary_input.toPlainText().strip(),
            "standards_used": self.standards_used_input.toPlainText().strip(),
            "stakeholder_summary": self.stakeholder_summary_input.toPlainText().strip(),
            "executive_summary": self.executive_summary_input.toPlainText().strip(),
            "technical_summary": self.technical_summary_input.toPlainText().strip(),
            "project_notes": self.project_notes_input.toPlainText().strip(),
            "conclusions": self.conclusions_input.toPlainText().strip(),
            "risk_rating": self.risk_rating_input.currentText().strip(),
        }

    def reset_project_form(self) -> None:
        self.current_edit_project_id = None

        self.form_title.setText("New Project")
        self.form_subtitle.setText(
            "Define the engagement, assets, summaries, and report content for a new assessment project."
        )

        self.project_name_input.clear()
        self.target_name_input.clear()
        self.platform_input.clear()
        self.client_name_input.clear()
        self.engagement_type_input.setCurrentIndex(0)
        self.assessment_start_input.clear()
        self.assessment_end_input.clear()
        self.risk_rating_input.setCurrentText("Not Rated")
        self.target_ips_input.clear()
        self.target_domains_input.clear()
        self.web_app_name_input.clear()
        self.mobile_app_name_input.clear()
        self.internal_app_name_input.clear()
        self.environment_name_input.clear()
        self.scope_summary_input.clear()
        self.out_of_scope_input.clear()
        self.attack_surface_input.clear()
        self.methodology_summary_input.clear()
        self.standards_used_input.clear()
        self.stakeholder_summary_input.clear()
        self.executive_summary_input.clear()
        self.technical_summary_input.clear()
        self.conclusions_input.clear()
        self.project_notes_input.clear()

        self.save_button.show()
        self.clear_form_button.show()
        self.update_button.hide()
        self.cancel_edit_button.hide()

    def go_to_new_project_form(self) -> None:
        self.reset_project_form()
        self.nav_list.setCurrentRow(1)

    def create_project_summary_card(self, project: Project) -> QFrame:
        card = ClickableProjectCard(project.id, self.set_current_project)
        card.setObjectName("ProjectItemCard")
        card.setProperty("selected", "true" if project.id == self.current_project_id else "false")
        card.style().unpolish(card)
        card.style().polish(card)

        layout = QVBoxLayout()
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(6)

        title = QLabel(project.project_name or "Untitled Project")
        title.setStyleSheet("font-size: 15px; font-weight: bold; color: #111827;")
        title.setWordWrap(True)

        meta = QLabel(
            f"{project.client_name or 'No client'} • "
            f"{project.target_name or 'No target'} • "
            f"{project.engagement_type or 'No type'}"
        )
        meta.setStyleSheet("font-size: 12px; color: #6b7280;")
        meta.setWordWrap(True)

        date_label = QLabel(self.format_project_date(project.created_at))
        date_label.setStyleSheet("font-size: 12px; color: #6b7280;")

        notes_preview = project.project_notes or project.executive_summary or project.scope_summary or ""
        notes_preview = notes_preview.strip()
        if len(notes_preview) > 180:
            notes_preview = notes_preview[:177] + "..."

        preview = QLabel(notes_preview or "No additional notes yet.")
        preview.setStyleSheet("font-size: 12px; color: #374151;")
        preview.setWordWrap(True)

        button_row = QHBoxLayout()
        button_row.setSpacing(8)

        edit_button = QPushButton("Edit")
        edit_button.setObjectName("SecondaryButton")
        edit_button.setFixedWidth(70)
        edit_button.clicked.connect(lambda _=False, pid=project.id: self.start_edit_project(pid))

        delete_button = QPushButton("Delete")
        delete_button.setObjectName("DangerButton")
        delete_button.setFixedWidth(70)
        delete_button.clicked.connect(lambda _=False, pid=project.id: self.delete_project(pid))

        button_row.addWidget(edit_button)
        button_row.addWidget(delete_button)
        button_row.addStretch()

        layout.addWidget(title)
        layout.addWidget(meta)
        layout.addWidget(date_label)
        layout.addWidget(preview)
        layout.addLayout(button_row)
        card.setLayout(layout)
        return card

    def handle_archive_filter_change(self, selected_text: str) -> None:
        if not selected_text:
            return
        self.selected_archive_filter = selected_text
        self.render_projects_browser()

    def load_projects(self) -> None:
        session = SessionLocal()
        try:
            self.all_projects_cache = (
                session.query(Project)
                .order_by(Project.created_at.desc(), Project.id.desc())
                .all()
            )
        finally:
            session.close()

        self.refresh_archive_groups()
        self.render_archive_list()
        self.render_projects_browser()
        self.render_dashboard_recent_projects()
        self.update_current_project_context()

    def refresh_archive_groups(self) -> None:
        groups = OrderedDict()
        groups["Recent"] = []
        groups["All Projects"] = []

        for project in self.all_projects_cache:
            groups["Recent"].append(project)
            groups["All Projects"].append(project)
            month_key = self.format_project_month_group(project.created_at)
            groups.setdefault(month_key, []).append(project)

        self.project_archive_groups = groups

    def render_archive_list(self) -> None:
        self.archive_list.blockSignals(True)
        self.archive_list.clear()

        for key in self.project_archive_groups.keys():
            self.archive_list.addItem(key)

        matches = self.archive_list.findItems(self.selected_archive_filter, Qt.MatchFlag.MatchExactly)
        if matches:
            self.archive_list.setCurrentItem(matches[0])
        elif self.archive_list.count() > 0:
            self.archive_list.setCurrentRow(0)
            self.selected_archive_filter = self.archive_list.item(0).text()

        self.archive_list.blockSignals(False)

    def render_projects_browser(self) -> None:
        while self.projects_container_layout.count():
            item = self.projects_container_layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()

        search_text = self.project_search_input.text().strip().lower()
        selected_group = self.selected_archive_filter or "Recent"
        projects = self.project_archive_groups.get(selected_group, [])

        filtered_projects = []
        for project in projects:
            haystack = " ".join(
                [
                    project.project_name or "",
                    project.client_name or "",
                    project.target_name or "",
                    project.platform or "",
                    project.engagement_type or "",
                    project.project_notes or "",
                    project.executive_summary or "",
                ]
            ).lower()

            if not search_text or search_text in haystack:
                filtered_projects.append(project)

        if not filtered_projects:
            empty_label = QLabel("No projects match the current search or archive filter.")
            empty_label.setStyleSheet("font-size: 13px; color: #6b7280;")
            empty_label.setWordWrap(True)
            self.projects_container_layout.addWidget(empty_label)
            self.projects_container_layout.addStretch()
            return

        for project in filtered_projects:
            self.projects_container_layout.addWidget(self.create_project_summary_card(project))

        self.projects_container_layout.addStretch()

    def render_dashboard_recent_projects(self) -> None:
        if self.recent_projects_container_layout is None:
            return

        while self.recent_projects_container_layout.count():
            item = self.recent_projects_container_layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()

        recent_projects = self.all_projects_cache[:5]
        if not recent_projects:
            label = QLabel("No recent projects yet. Start by creating your first project.")
            label.setStyleSheet("font-size: 13px; color: #6b7280;")
            label.setWordWrap(True)
            self.recent_projects_container_layout.addWidget(label)
            return

        for project in recent_projects:
            self.recent_projects_container_layout.addWidget(self.create_project_summary_card(project))

    def set_current_project(self, project_id: int) -> None:
        self.current_project_id = project_id
        self.current_selected_finding_id = None
        self.update_current_project_context()
        self.load_projects()
        self.load_findings()
        self.load_reference_finding_list()
        self.load_references()
        self.update_checklist_context()
        self.update_export_context()

    def update_current_project_context(self) -> None:
        session = SessionLocal()
        try:
            project = None
            if self.current_project_id is not None:
                project = session.get(Project, self.current_project_id)

            if project is None:
                self.current_project_name_label.setText("No project selected")
                self.current_project_date_label.setText("")
                self.current_project_meta_label.setText(
                    "Select a saved project to make it the active workspace."
                )
                return

            findings_count = session.query(Finding).filter(Finding.project_id == project.id).count()

            self.current_project_name_label.setText(project.project_name)
            self.current_project_date_label.setText(self.format_project_date(project.created_at))
            self.current_project_meta_label.setText(
                f"Client: {project.client_name or 'Not specified'}   •   "
                f"Target: {project.target_name or 'Not specified'}   •   "
                f"Type: {project.engagement_type or 'Not specified'}   •   "
                f"Risk: {project.risk_rating or 'Not Rated'}   •   "
                f"Findings: {findings_count}"
            )
        finally:
            session.close()

    def handle_save_project(self) -> None:
        data = self.collect_project_form_data()
        if not data["project_name"]:
            QMessageBox.warning(self, "Missing Project Name", "Please enter a project name before saving.")
            return

        session = SessionLocal()
        try:
            project = Project(**data)
            session.add(project)
            session.commit()
            session.refresh(project)

            self.current_project_id = project.id
            self.selected_archive_filter = "Recent"

            QMessageBox.information(
                self,
                "Project Saved",
                f'Project "{data["project_name"]}" was saved successfully.',
            )
        finally:
            session.close()

        self.reset_project_form()
        self.load_projects()
        self.load_findings()
        self.load_reference_finding_list()
        self.load_references()
        self.update_checklist_context()
        self.update_export_context()

    def start_edit_project(self, project_id: int) -> None:
        session = SessionLocal()
        try:
            project = session.get(Project, project_id)
            if project is None:
                QMessageBox.warning(self, "Project Not Found", "The selected project could not be loaded.")
                return

            self.current_edit_project_id = project.id

            self.project_name_input.setText(project.project_name or "")
            self.target_name_input.setText(project.target_name or "")
            self.platform_input.setText(project.platform or "")
            self.client_name_input.setText(project.client_name or "")
            self.engagement_type_input.setCurrentText(project.engagement_type or "Other")
            self.assessment_start_input.setText(project.assessment_start or "")
            self.assessment_end_input.setText(project.assessment_end or "")
            self.risk_rating_input.setCurrentText(project.risk_rating or "Not Rated")
            self.target_ips_input.setPlainText(project.target_ips or "")
            self.target_domains_input.setPlainText(project.target_domains or "")
            self.web_app_name_input.setText(project.web_app_name or "")
            self.mobile_app_name_input.setText(project.mobile_app_name or "")
            self.internal_app_name_input.setText(project.internal_app_name or "")
            self.environment_name_input.setText(project.environment_name or "")
            self.scope_summary_input.setPlainText(project.scope_summary or "")
            self.out_of_scope_input.setPlainText(project.out_of_scope or "")
            self.attack_surface_input.setPlainText(project.attack_surface or "")
            self.methodology_summary_input.setPlainText(project.methodology_summary or "")
            self.standards_used_input.setPlainText(project.standards_used or "")
            self.stakeholder_summary_input.setPlainText(project.stakeholder_summary or "")
            self.executive_summary_input.setPlainText(project.executive_summary or "")
            self.technical_summary_input.setPlainText(project.technical_summary or "")
            self.conclusions_input.setPlainText(project.conclusions or "")
            self.project_notes_input.setPlainText(project.project_notes or "")

            self.form_title.setText("Edit Project")
            self.form_subtitle.setText("Update the selected engagement record.")

            self.save_button.hide()
            self.clear_form_button.hide()
            self.update_button.show()
            self.cancel_edit_button.show()

            self.nav_list.setCurrentRow(1)
        finally:
            session.close()

    def handle_update_project(self) -> None:
        if self.current_edit_project_id is None:
            return

        data = self.collect_project_form_data()
        if not data["project_name"]:
            QMessageBox.warning(self, "Missing Project Name", "Please enter a project name before updating.")
            return

        session = SessionLocal()
        try:
            project = session.get(Project, self.current_edit_project_id)
            if project is None:
                QMessageBox.warning(self, "Project Not Found", "The selected project could not be updated.")
                return

            for key, value in data.items():
                setattr(project, key, value)

            session.commit()

            if self.current_project_id == project.id:
                self.current_project_id = project.id

            QMessageBox.information(
                self,
                "Project Updated",
                f'Project "{data["project_name"]}" was updated successfully.',
            )
        finally:
            session.close()

        self.reset_project_form()
        self.load_projects()
        self.update_current_project_context()
        self.update_checklist_context()
        self.update_export_context()

    def delete_project(self, project_id: int) -> None:
        confirm = QMessageBox.question(
            self,
            "Delete Project",
            "Delete this project and all its findings, references, and checklist state?",
        )
        if confirm != QMessageBox.StandardButton.Yes:
            return

        session = SessionLocal()
        try:
            finding_ids = [
                row[0]
                for row in session.query(Finding.id).filter(Finding.project_id == project_id).all()
            ]
            if finding_ids:
                session.query(Reference).filter(Reference.finding_id.in_(finding_ids)).delete(
                    synchronize_session=False
                )
            session.query(Finding).filter(Finding.project_id == project_id).delete(synchronize_session=False)
            session.query(ChecklistItem).filter(ChecklistItem.project_id == project_id).delete(
                synchronize_session=False
            )
            session.query(Project).filter(Project.id == project_id).delete(synchronize_session=False)
            session.commit()
        finally:
            session.close()

        if self.current_project_id == project_id:
            self.current_project_id = None
            self.current_selected_finding_id = None

        self.reset_project_form()
        self.reset_finding_form()
        self.reset_reference_form()
        self.load_projects()
        self.load_findings()
        self.load_reference_finding_list()
        self.load_references()
        self.update_checklist_context()
        self.update_export_context()

    def build_findings_page(self) -> QWidget:
        page = QWidget()

        page_scroll = QScrollArea()
        page_scroll.setWidgetResizable(True)
        page_scroll.setFrameShape(QFrame.Shape.NoFrame)

        page_content = QWidget()
        outer_layout = QVBoxLayout()
        outer_layout.setContentsMargins(0, 0, 0, 0)
        outer_layout.setSpacing(16)

        page_title = QLabel("Findings")
        page_title.setStyleSheet("font-size: 30px; font-weight: bold;")

        self.findings_project_context_label = QLabel(
            "Select a current project to manage findings."
        )
        self.findings_project_context_label.setStyleSheet(
            "font-size: 14px; color: #4b5563;"
        )

        self.findings_project_card = QFrame()
        self.findings_project_card.setObjectName("Card")

        findings_project_layout = QVBoxLayout()
        findings_project_layout.setContentsMargins(24, 20, 24, 20)
        findings_project_layout.setSpacing(6)

        self.findings_current_project_name = QLabel("No active project")
        self.findings_current_project_name.setStyleSheet(
            "font-size: 18px; font-weight: bold; color: #111827;"
        )

        self.findings_current_project_meta = QLabel(
            "Go to Projects and select a current project first."
        )
        self.findings_current_project_meta.setWordWrap(True)
        self.findings_current_project_meta.setStyleSheet(
            "font-size: 13px; color: #6b7280;"
        )

        findings_project_layout.addWidget(self.findings_current_project_name)
        findings_project_layout.addWidget(self.findings_current_project_meta)
        self.findings_project_card.setLayout(findings_project_layout)

        content_row = QHBoxLayout()
        content_row.setSpacing(10)
        content_row.setAlignment(Qt.AlignmentFlag.AlignTop)

        form_card = QFrame()
        form_card.setObjectName("Card")
        form_card.setMinimumWidth(540)
        form_card.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Maximum)

        form_layout = QVBoxLayout()
        form_layout.setContentsMargins(24, 24, 24, 24)
        form_layout.setSpacing(10)

        self.finding_form_title = QLabel("New Finding")
        self.finding_form_title.setStyleSheet("font-size: 20px; font-weight: bold;")

        self.finding_form_subtitle = QLabel(
            "Document a finding for the currently selected project."
        )
        self.finding_form_subtitle.setStyleSheet("font-size: 13px; color: #6b7280;")

        self.finding_title_input = QLineEdit()
        self.finding_title_input.setPlaceholderText("Example: Weak SSH Credentials")

        self.finding_template_input = QComboBox()
        self.finding_template_input.addItem("Select a template...")
        self.finding_template_input.addItems(get_template_names())

        apply_template_button = QPushButton("Apply Template")
        apply_template_button.setObjectName("SecondaryButton")
        apply_template_button.setFixedWidth(130)
        apply_template_button.clicked.connect(self.handle_apply_finding_template)

        self.finding_severity_input = QComboBox()
        self.finding_severity_input.addItems(["Critical", "High", "Medium", "Low", "Info"])
        self.finding_severity_input.setCurrentText("Info")

        self.finding_description_input = self.create_multiline_input(
            "Describe the issue clearly and professionally.",
            110,
        )
        self.finding_evidence_input = self.create_multiline_input(
            "Add commands, output, observations, or proof of concept notes.",
            100,
        )
        self.finding_remediation_input = self.create_multiline_input(
            "Explain how the issue should be fixed or mitigated.",
            100,
        )

        self.finding_attachments_list = QListWidget()
        self.finding_attachments_list.setObjectName("AttachmentList")
        self.finding_attachments_list.setFixedHeight(110)

        attachment_buttons_row = QHBoxLayout()
        attachment_buttons_row.setSpacing(8)

        add_attachment_button = QPushButton("Add Files")
        add_attachment_button.setObjectName("SecondaryButton")
        add_attachment_button.setFixedWidth(110)
        add_attachment_button.clicked.connect(self.add_finding_attachments)

        remove_attachment_button = QPushButton("Remove Selected")
        remove_attachment_button.setObjectName("SecondaryButton")
        remove_attachment_button.setFixedWidth(140)
        remove_attachment_button.clicked.connect(self.remove_selected_finding_attachment)

        attachment_buttons_row.addWidget(add_attachment_button)
        attachment_buttons_row.addWidget(remove_attachment_button)
        attachment_buttons_row.addStretch()

        self.save_finding_button = QPushButton("Save Finding")
        self.save_finding_button.clicked.connect(self.handle_save_finding)
        self.save_finding_button.setFixedWidth(140)

        self.clear_finding_form_button = QPushButton("Clear Form")
        self.clear_finding_form_button.clicked.connect(self.reset_finding_form)
        self.clear_finding_form_button.setObjectName("SecondaryButton")
        self.clear_finding_form_button.setFixedWidth(120)

        self.update_finding_button = QPushButton("Update Finding")
        self.update_finding_button.clicked.connect(self.handle_update_finding)
        self.update_finding_button.setFixedWidth(150)
        self.update_finding_button.hide()

        self.cancel_finding_edit_button = QPushButton("Cancel Editing")
        self.cancel_finding_edit_button.clicked.connect(self.reset_finding_form)
        self.cancel_finding_edit_button.setObjectName("SecondaryButton")
        self.cancel_finding_edit_button.setFixedWidth(130)
        self.cancel_finding_edit_button.hide()

        finding_button_row = QHBoxLayout()
        finding_button_row.setSpacing(10)
        finding_button_row.addWidget(self.save_finding_button)
        finding_button_row.addWidget(self.clear_finding_form_button)
        finding_button_row.addWidget(self.update_finding_button)
        finding_button_row.addWidget(self.cancel_finding_edit_button)
        finding_button_row.addStretch()

        form_layout.addWidget(self.finding_form_title)
        form_layout.addWidget(self.finding_form_subtitle)
        form_layout.addSpacing(8)
        form_layout.addWidget(self.create_input_label("Finding Title"))
        form_layout.addWidget(self.finding_title_input)
        form_layout.addWidget(self.create_input_label("Finding Template"))
        form_layout.addWidget(self.finding_template_input)
        form_layout.addWidget(apply_template_button, alignment=Qt.AlignmentFlag.AlignLeft)
        form_layout.addWidget(self.create_input_label("Severity"))
        form_layout.addWidget(self.finding_severity_input)
        form_layout.addWidget(self.create_input_label("Description"))
        form_layout.addWidget(self.finding_description_input)
        form_layout.addWidget(self.create_input_label("Evidence"))
        form_layout.addWidget(self.finding_evidence_input)
        form_layout.addWidget(self.create_input_label("Evidence Attachments"))
        form_layout.addWidget(self.finding_attachments_list)
        form_layout.addLayout(attachment_buttons_row)
        form_layout.addWidget(self.create_input_label("Remediation"))
        form_layout.addWidget(self.finding_remediation_input)
        form_layout.addSpacing(8)
        form_layout.addLayout(finding_button_row)
        form_card.setLayout(form_layout)

        findings_card = QFrame()
        findings_card.setObjectName("Card")
        findings_card.setMinimumWidth(620)
        findings_card.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Expanding)

        findings_layout = QVBoxLayout()
        findings_layout.setContentsMargins(24, 24, 24, 24)
        findings_layout.setSpacing(12)

        findings_title = QLabel("Project Findings")
        findings_title.setStyleSheet("font-size: 20px; font-weight: bold;")

        findings_subtitle = QLabel(
            "Review and maintain the findings recorded for the active project."
        )
        findings_subtitle.setStyleSheet("font-size: 13px; color: #6b7280;")

        findings_scroll_area = QScrollArea()
        findings_scroll_area.setWidgetResizable(True)
        findings_scroll_area.setFrameShape(QFrame.Shape.NoFrame)

        self.findings_container = QWidget()
        self.findings_container_layout = QVBoxLayout()
        self.findings_container_layout.setContentsMargins(0, 0, 0, 0)
        self.findings_container_layout.setSpacing(14)
        self.findings_container_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        self.findings_container.setLayout(self.findings_container_layout)
        findings_scroll_area.setWidget(self.findings_container)

        findings_layout.addWidget(findings_title)
        findings_layout.addWidget(findings_subtitle)
        findings_layout.addWidget(findings_scroll_area)
        findings_card.setLayout(findings_layout)

        content_row.addWidget(form_card, 1)
        content_row.addWidget(findings_card, 1)

        outer_layout.addWidget(page_title)
        outer_layout.addWidget(self.findings_project_context_label)
        outer_layout.addWidget(self.findings_project_card)
        outer_layout.addLayout(content_row)
        outer_layout.addStretch()

        page_content.setLayout(outer_layout)
        page_scroll.setWidget(page_content)

        wrapper_layout = QVBoxLayout()
        wrapper_layout.setContentsMargins(0, 0, 0, 0)
        wrapper_layout.addWidget(page_scroll)

        page.setLayout(wrapper_layout)
        return page

    def update_findings_project_context(self) -> None:
        session = SessionLocal()
        try:
            project = session.get(Project, self.current_project_id) if self.current_project_id else None
            if project is None:
                self.findings_current_project_name.setText("No active project")
                self.findings_current_project_meta.setText("Go to Projects and select a current project first.")
                return

            findings_count = session.query(Finding).filter(Finding.project_id == project.id).count()
            self.findings_current_project_name.setText(project.project_name)
            self.findings_current_project_meta.setText(
                f"Client: {project.client_name or 'Not specified'}   •   "
                f"Target: {project.target_name or 'Not specified'}   •   "
                f"Findings: {findings_count}"
            )
        finally:
            session.close()

    def reset_finding_form(self) -> None:
        self.current_edit_finding_id = None
        self.current_finding_attachments = []

        self.finding_form_title.setText("New Finding")
        self.finding_form_subtitle.setText("Document a finding for the currently selected project.")
        self.finding_title_input.clear()
        self.finding_template_input.setCurrentIndex(0)
        self.finding_severity_input.setCurrentText("Info")
        self.finding_description_input.clear()
        self.finding_evidence_input.clear()
        self.finding_remediation_input.clear()
        self.refresh_finding_attachment_list()

        self.save_finding_button.show()
        self.clear_finding_form_button.show()
        self.update_finding_button.hide()
        self.cancel_finding_edit_button.hide()

    def refresh_finding_attachment_list(self) -> None:
        self.finding_attachments_list.clear()
        for attachment in self.current_finding_attachments:
            self.finding_attachments_list.addItem(attachment)

    def add_finding_attachments(self) -> None:
        files, _ = QFileDialog.getOpenFileNames(self, "Select Evidence Files")
        if not files:
            return
        for file_path in files:
            if file_path not in self.current_finding_attachments:
                self.current_finding_attachments.append(file_path)
        self.refresh_finding_attachment_list()

    def remove_selected_finding_attachment(self) -> None:
        item = self.finding_attachments_list.currentItem()
        if item is None:
            return
        path_text = item.text()
        self.current_finding_attachments = [p for p in self.current_finding_attachments if p != path_text]
        self.refresh_finding_attachment_list()

    def handle_apply_finding_template(self) -> None:
        template_name = self.finding_template_input.currentText().strip()
        if not template_name or template_name == "Select a template...":
            return

        template_data = get_template_data(template_name)
        if not template_data:
            return

        self.finding_title_input.setText(template_name)
        self.finding_severity_input.setCurrentText(template_data.get("severity", "Info"))
        self.finding_description_input.setPlainText(template_data.get("description", ""))
        self.finding_evidence_input.setPlainText(template_data.get("evidence", ""))
        self.finding_remediation_input.setPlainText(template_data.get("remediation", ""))

    def load_findings(self) -> None:
        self.update_findings_project_context()

        while self.findings_container_layout.count():
            item = self.findings_container_layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()

        if self.current_project_id is None:
            label = QLabel("No active project selected.")
            label.setStyleSheet("font-size: 13px; color: #6b7280;")
            self.findings_container_layout.addWidget(label)
            self.findings_container_layout.addStretch()
            return

        session = SessionLocal()
        try:
            findings = (
                session.query(Finding)
                .filter(Finding.project_id == self.current_project_id)
                .order_by(Finding.id.asc())
                .all()
            )
        finally:
            session.close()

        if not findings:
            label = QLabel("No findings recorded yet for this project.")
            label.setStyleSheet("font-size: 13px; color: #6b7280;")
            self.findings_container_layout.addWidget(label)
            self.findings_container_layout.addStretch()
            return

        for finding in findings:
            card = QFrame()
            card.setObjectName("Card")
            layout = QVBoxLayout()
            layout.setContentsMargins(18, 18, 18, 18)
            layout.setSpacing(8)

            top_row = QHBoxLayout()
            title = QLabel(finding.title)
            title.setStyleSheet("font-size: 16px; font-weight: bold; color: #111827;")
            title.setWordWrap(True)

            severity = QLabel((finding.severity or "Info").strip().capitalize())
            severity.setObjectName(self.severity_object_name(finding.severity or "Info"))
            top_row.addWidget(title, 1)
            top_row.addWidget(severity, 0, Qt.AlignmentFlag.AlignRight)

            description = finding.description or ""
            if len(description) > 220:
                description = description[:217] + "..."
            desc_label = QLabel(description or "No description provided.")
            desc_label.setStyleSheet("font-size: 13px; color: #374151;")
            desc_label.setWordWrap(True)

            attachments = finding.get_attachments()
            attach_text = f"Attachments: {len(attachments)}"
            attachments_label = QLabel(attach_text)
            attachments_label.setStyleSheet("font-size: 12px; color: #6b7280;")

            button_row = QHBoxLayout()
            edit_button = QPushButton("Edit")
            edit_button.setObjectName("SecondaryButton")
            edit_button.setFixedWidth(80)
            edit_button.clicked.connect(lambda _=False, fid=finding.id: self.start_edit_finding(fid))

            delete_button = QPushButton("Delete")
            delete_button.setObjectName("DangerButton")
            delete_button.setFixedWidth(80)
            delete_button.clicked.connect(lambda _=False, fid=finding.id: self.delete_finding(fid))

            button_row.addWidget(edit_button)
            button_row.addWidget(delete_button)
            button_row.addStretch()

            layout.addLayout(top_row)
            layout.addWidget(desc_label)
            layout.addWidget(attachments_label)
            layout.addLayout(button_row)
            card.setLayout(layout)
            self.findings_container_layout.addWidget(card)

        self.findings_container_layout.addStretch()

    def handle_save_finding(self) -> None:
        if self.current_project_id is None:
            QMessageBox.warning(self, "No Active Project", "Please select a current project first.")
            return

        title = self.finding_title_input.text().strip()
        if not title:
            QMessageBox.warning(self, "Missing Finding Title", "Please enter a finding title.")
            return

        session = SessionLocal()
        try:
            finding = Finding(
                project_id=self.current_project_id,
                title=title,
                severity=self.finding_severity_input.currentText().strip(),
                description=self.finding_description_input.toPlainText().strip(),
                evidence=self.finding_evidence_input.toPlainText().strip(),
                remediation=self.finding_remediation_input.toPlainText().strip(),
            )
            session.add(finding)
            session.commit()
            session.refresh(finding)

            copied_paths = copy_evidence_files(
                self.current_finding_attachments,
                self.current_project_id,
                finding.id,
            )
            finding.set_attachments(copied_paths)
            session.commit()
        finally:
            session.close()

        QMessageBox.information(self, "Finding Saved", f'Finding "{title}" was saved successfully.')
        self.reset_finding_form()
        self.load_findings()
        self.load_reference_finding_list()
        self.load_references()
        self.update_checklist_context()
        self.update_export_context()

    def start_edit_finding(self, finding_id: int) -> None:
        session = SessionLocal()
        try:
            finding = session.get(Finding, finding_id)
            if finding is None:
                return

            self.current_edit_finding_id = finding.id
            self.finding_title_input.setText(finding.title or "")
            self.finding_severity_input.setCurrentText(finding.severity or "Info")
            self.finding_description_input.setPlainText(finding.description or "")
            self.finding_evidence_input.setPlainText(finding.evidence or "")
            self.finding_remediation_input.setPlainText(finding.remediation or "")
            self.current_finding_attachments = finding.get_attachments()
            self.refresh_finding_attachment_list()

            self.finding_form_title.setText("Edit Finding")
            self.finding_form_subtitle.setText("Update the selected finding for the active project.")

            self.save_finding_button.hide()
            self.clear_finding_form_button.hide()
            self.update_finding_button.show()
            self.cancel_finding_edit_button.show()

            self.nav_list.setCurrentRow(2)
        finally:
            session.close()

    def handle_update_finding(self) -> None:
        if self.current_edit_finding_id is None:
            return

        session = SessionLocal()
        try:
            finding = session.get(Finding, self.current_edit_finding_id)
            if finding is None:
                return

            finding.title = self.finding_title_input.text().strip()
            finding.severity = self.finding_severity_input.currentText().strip()
            finding.description = self.finding_description_input.toPlainText().strip()
            finding.evidence = self.finding_evidence_input.toPlainText().strip()
            finding.remediation = self.finding_remediation_input.toPlainText().strip()

            existing_attachments = finding.get_attachments()
            incoming = []
            preserved = []

            for path_text in self.current_finding_attachments:
                if path_text.startswith("reportforge_data") or "/reportforge_data/" in path_text.replace("\\", "/"):
                    preserved.append(path_text)
                elif path_text in existing_attachments:
                    preserved.append(path_text)
                else:
                    incoming.append(path_text)

            copied_paths = copy_evidence_files(incoming, finding.project_id, finding.id)
            finding.set_attachments(preserved + copied_paths)
            session.commit()
        finally:
            session.close()

        QMessageBox.information(self, "Finding Updated", "Finding updated successfully.")
        self.reset_finding_form()
        self.load_findings()
        self.load_references()
        self.update_checklist_context()
        self.update_export_context()

    def delete_finding(self, finding_id: int) -> None:
        confirm = QMessageBox.question(self, "Delete Finding", "Delete this finding and its references?")
        if confirm != QMessageBox.StandardButton.Yes:
            return

        session = SessionLocal()
        try:
            session.query(Reference).filter(Reference.finding_id == finding_id).delete(synchronize_session=False)
            session.query(Finding).filter(Finding.id == finding_id).delete(synchronize_session=False)
            session.commit()
        finally:
            session.close()

        if self.current_selected_finding_id == finding_id:
            self.current_selected_finding_id = None

        self.reset_finding_form()
        self.load_findings()
        self.load_reference_finding_list()
        self.load_references()
        self.update_checklist_context()
        self.update_export_context()

    def build_references_page(self) -> QWidget:
        page = QWidget()

        page_scroll = QScrollArea()
        page_scroll.setWidgetResizable(True)
        page_scroll.setFrameShape(QFrame.Shape.NoFrame)

        page_content = QWidget()
        outer_layout = QVBoxLayout()
        outer_layout.setContentsMargins(0, 0, 0, 0)
        outer_layout.setSpacing(16)

        page_title = QLabel("References")
        page_title.setStyleSheet("font-size: 30px; font-weight: bold;")

        self.references_context_label = QLabel(
            "Select a project and a finding to manage reference links."
        )
        self.references_context_label.setStyleSheet("font-size: 14px; color: #4b5563;")

        top_row = QHBoxLayout()
        top_row.setSpacing(16)
        top_row.setAlignment(Qt.AlignmentFlag.AlignTop)

        finding_selector_card = QFrame()
        finding_selector_card.setObjectName("Card")
        finding_selector_card.setMinimumWidth(440)
        finding_selector_card.setMaximumWidth(520)

        finding_selector_layout = QVBoxLayout()
        finding_selector_layout.setContentsMargins(24, 24, 24, 24)
        finding_selector_layout.setSpacing(12)

        finding_selector_title = QLabel("Finding Selection")
        finding_selector_title.setStyleSheet("font-size: 20px; font-weight: bold;")

        self.reference_project_label = QLabel("No active project")
        self.reference_project_label.setStyleSheet("font-size: 13px; color: #6b7280;")

        self.reference_finding_list = QListWidget()
        self.reference_finding_list.setObjectName("FindingSelectionList")
        self.reference_finding_list.setMinimumHeight(180)
        self.reference_finding_list.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.reference_finding_list.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        self.reference_finding_list.setTextElideMode(Qt.TextElideMode.ElideNone)
        self.reference_finding_list.setWordWrap(False)
        self.reference_finding_list.currentRowChanged.connect(self.handle_reference_finding_selection)

        finding_selector_layout.addWidget(finding_selector_title)
        finding_selector_layout.addWidget(self.reference_project_label)
        finding_selector_layout.addWidget(self.reference_finding_list)
        finding_selector_card.setLayout(finding_selector_layout)

        selected_finding_card = QFrame()
        selected_finding_card.setObjectName("Card")

        selected_finding_layout = QVBoxLayout()
        selected_finding_layout.setContentsMargins(24, 20, 24, 20)
        selected_finding_layout.setSpacing(6)

        selected_finding_title = QLabel("Selected Finding")
        selected_finding_title.setStyleSheet("font-size: 18px; font-weight: bold;")

        self.selected_reference_finding_name = QLabel("No finding selected")
        self.selected_reference_finding_name.setStyleSheet(
            "font-size: 16px; font-weight: bold; color: #111827;"
        )

        self.selected_reference_finding_meta = QLabel(
            "Choose a finding from the list to manage its references."
        )
        self.selected_reference_finding_meta.setWordWrap(True)
        self.selected_reference_finding_meta.setStyleSheet("font-size: 13px; color: #6b7280;")

        selected_finding_layout.addWidget(selected_finding_title)
        selected_finding_layout.addWidget(self.selected_reference_finding_name)
        selected_finding_layout.addWidget(self.selected_reference_finding_meta)
        selected_finding_card.setLayout(selected_finding_layout)

        top_row.addWidget(finding_selector_card)
        top_row.addWidget(selected_finding_card, 1)

        content_row = QHBoxLayout()
        content_row.setSpacing(16)
        content_row.setAlignment(Qt.AlignmentFlag.AlignTop)

        reference_form_card = QFrame()
        reference_form_card.setObjectName("Card")
        reference_form_card.setMinimumWidth(540)
        reference_form_card.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Maximum)

        reference_form_layout = QVBoxLayout()
        reference_form_layout.setContentsMargins(24, 24, 24, 24)
        reference_form_layout.setSpacing(10)

        self.reference_form_title = QLabel("New Reference")
        self.reference_form_title.setStyleSheet("font-size: 20px; font-weight: bold;")

        self.reference_form_subtitle = QLabel(
            "Add a link or supporting source for the selected finding."
        )
        self.reference_form_subtitle.setStyleSheet("font-size: 13px; color: #6b7280;")

        self.reference_title_input = QLineEdit()
        self.reference_title_input.setPlaceholderText("Example: CVE-2021-41773")

        self.reference_type_input = QComboBox()
        self.reference_type_input.addItems(
            ["CVE", "CWE", "NVD", "OWASP", "Vendor Advisory", "Blog", "Custom"]
        )
        self.reference_type_input.setCurrentText("Custom")

        self.reference_url_input = QLineEdit()
        self.reference_url_input.setPlaceholderText("https://...")

        helper_buttons_row = QHBoxLayout()
        helper_buttons_row.setSpacing(8)

        auto_fill_button = QPushButton("Auto Fill from Title")
        auto_fill_button.setObjectName("SecondaryButton")
        auto_fill_button.setFixedWidth(150)
        auto_fill_button.clicked.connect(self.handle_reference_auto_fill)

        normalize_url_button = QPushButton("Normalize URL")
        normalize_url_button.setObjectName("SecondaryButton")
        normalize_url_button.setFixedWidth(120)
        normalize_url_button.clicked.connect(self.handle_reference_normalize_url)

        helper_buttons_row.addWidget(auto_fill_button)
        helper_buttons_row.addWidget(normalize_url_button)
        helper_buttons_row.addStretch()

        self.reference_note_input = self.create_multiline_input(
            "Add an optional note about why this reference matters.",
            120,
        )

        self.save_reference_button = QPushButton("Save Reference")
        self.save_reference_button.clicked.connect(self.handle_save_reference)
        self.save_reference_button.setFixedWidth(150)

        self.clear_reference_form_button = QPushButton("Clear Form")
        self.clear_reference_form_button.clicked.connect(self.reset_reference_form)
        self.clear_reference_form_button.setObjectName("SecondaryButton")
        self.clear_reference_form_button.setFixedWidth(120)

        self.update_reference_button = QPushButton("Update Reference")
        self.update_reference_button.clicked.connect(self.handle_update_reference)
        self.update_reference_button.setFixedWidth(160)
        self.update_reference_button.hide()

        self.cancel_reference_edit_button = QPushButton("Cancel Editing")
        self.cancel_reference_edit_button.clicked.connect(self.reset_reference_form)
        self.cancel_reference_edit_button.setObjectName("SecondaryButton")
        self.cancel_reference_edit_button.setFixedWidth(130)
        self.cancel_reference_edit_button.hide()

        reference_button_row = QHBoxLayout()
        reference_button_row.setSpacing(10)
        reference_button_row.addWidget(self.save_reference_button)
        reference_button_row.addWidget(self.clear_reference_form_button)
        reference_button_row.addWidget(self.update_reference_button)
        reference_button_row.addWidget(self.cancel_reference_edit_button)
        reference_button_row.addStretch()

        reference_form_layout.addWidget(self.reference_form_title)
        reference_form_layout.addWidget(self.reference_form_subtitle)
        reference_form_layout.addSpacing(8)
        reference_form_layout.addWidget(self.create_input_label("Reference Title"))
        reference_form_layout.addWidget(self.reference_title_input)
        reference_form_layout.addWidget(self.create_input_label("Reference Type"))
        reference_form_layout.addWidget(self.reference_type_input)
        reference_form_layout.addWidget(self.create_input_label("URL"))
        reference_form_layout.addWidget(self.reference_url_input)
        reference_form_layout.addLayout(helper_buttons_row)
        reference_form_layout.addWidget(self.create_input_label("Notes"))
        reference_form_layout.addWidget(self.reference_note_input)
        reference_form_layout.addSpacing(8)
        reference_form_layout.addLayout(reference_button_row)
        reference_form_card.setLayout(reference_form_layout)

        references_list_card = QFrame()
        references_list_card.setObjectName("Card")
        references_list_card.setMinimumWidth(620)
        references_list_card.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Expanding)

        references_list_layout = QVBoxLayout()
        references_list_layout.setContentsMargins(24, 24, 24, 24)
        references_list_layout.setSpacing(12)

        references_list_title = QLabel("Finding References")
        references_list_title.setStyleSheet("font-size: 20px; font-weight: bold;")

        references_list_subtitle = QLabel(
            "Open, edit, and manage the sources attached to the selected finding."
        )
        references_list_subtitle.setStyleSheet("font-size: 13px; color: #6b7280;")

        references_scroll_area = QScrollArea()
        references_scroll_area.setWidgetResizable(True)
        references_scroll_area.setFrameShape(QFrame.Shape.NoFrame)

        self.references_container = QWidget()
        self.references_container_layout = QVBoxLayout()
        self.references_container_layout.setContentsMargins(0, 0, 0, 0)
        self.references_container_layout.setSpacing(14)
        self.references_container_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        self.references_container.setLayout(self.references_container_layout)
        references_scroll_area.setWidget(self.references_container)

        references_list_layout.addWidget(references_list_title)
        references_list_layout.addWidget(references_list_subtitle)
        references_list_layout.addWidget(references_scroll_area)
        references_list_card.setLayout(references_list_layout)

        content_row.addWidget(reference_form_card, 1)
        content_row.addWidget(references_list_card, 1)

        outer_layout.addWidget(page_title)
        outer_layout.addWidget(self.references_context_label)
        outer_layout.addLayout(top_row)
        outer_layout.addLayout(content_row)
        outer_layout.addStretch()

        page_content.setLayout(outer_layout)
        page_scroll.setWidget(page_content)

        wrapper_layout = QVBoxLayout()
        wrapper_layout.setContentsMargins(0, 0, 0, 0)
        wrapper_layout.addWidget(page_scroll)

        page.setLayout(wrapper_layout)
        return page

    def reset_reference_form(self) -> None:
        self.current_edit_reference_id = None
        self.reference_form_title.setText("New Reference")
        self.reference_form_subtitle.setText("Add a link or supporting source for the selected finding.")
        self.reference_title_input.clear()
        self.reference_type_input.setCurrentText("Custom")
        self.reference_url_input.clear()
        self.reference_note_input.clear()
        self.save_reference_button.show()
        self.clear_reference_form_button.show()
        self.update_reference_button.hide()
        self.cancel_reference_edit_button.hide()

    def load_reference_finding_list(self) -> None:
        self.reference_finding_list.clear()
        self.reference_project_label.setText("No active project" if self.current_project_id is None else "Active project selected")
        self.selected_reference_finding_name.setText("No finding selected")
        self.selected_reference_finding_meta.setText("Choose a finding from the list to manage its references.")

        if self.current_project_id is None:
            return

        session = SessionLocal()
        try:
            project = session.get(Project, self.current_project_id)
            findings = (
                session.query(Finding)
                .filter(Finding.project_id == self.current_project_id)
                .order_by(Finding.id.asc())
                .all()
            )
            if project is not None:
                self.reference_project_label.setText(
                    f"Project: {project.project_name}   •   Findings: {len(findings)}"
                )

            for finding in findings:
                item = QListWidgetItem(finding.title)
                item.setData(Qt.ItemDataRole.UserRole, finding.id)
                self.reference_finding_list.addItem(item)

            if self.current_selected_finding_id is not None:
                for index in range(self.reference_finding_list.count()):
                    item = self.reference_finding_list.item(index)
                    if item.data(Qt.ItemDataRole.UserRole) == self.current_selected_finding_id:
                        self.reference_finding_list.setCurrentRow(index)
                        break
        finally:
            session.close()

    def handle_reference_finding_selection(self, row: int) -> None:
        if row < 0:
            self.current_selected_finding_id = None
            self.load_references()
            return

        item = self.reference_finding_list.item(row)
        if item is None:
            return

        self.current_selected_finding_id = item.data(Qt.ItemDataRole.UserRole)

        session = SessionLocal()
        try:
            finding = session.get(Finding, self.current_selected_finding_id)
            if finding is None:
                return
            self.selected_reference_finding_name.setText(finding.title)
            self.selected_reference_finding_meta.setText(
                f"Severity: {finding.severity or 'Info'}   •   References can be attached here."
            )
        finally:
            session.close()

        self.load_references()

    def load_references(self) -> None:
        while self.references_container_layout.count():
            item = self.references_container_layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()

        if self.current_selected_finding_id is None:
            label = QLabel("Select a finding to view or manage references.")
            label.setStyleSheet("font-size: 13px; color: #6b7280;")
            self.references_container_layout.addWidget(label)
            self.references_container_layout.addStretch()
            return

        session = SessionLocal()
        try:
            references = (
                session.query(Reference)
                .filter(Reference.finding_id == self.current_selected_finding_id)
                .order_by(Reference.id.asc())
                .all()
            )
        finally:
            session.close()

        if not references:
            label = QLabel("No references attached to the selected finding.")
            label.setStyleSheet("font-size: 13px; color: #6b7280;")
            self.references_container_layout.addWidget(label)
            self.references_container_layout.addStretch()
            return

        for reference in references:
            card = QFrame()
            card.setObjectName("Card")
            layout = QVBoxLayout()
            layout.setContentsMargins(18, 18, 18, 18)
            layout.setSpacing(8)

            title_row = QHBoxLayout()

            title = QLabel(reference.title)
            title.setStyleSheet("font-size: 15px; font-weight: bold; color: #111827;")
            title.setWordWrap(True)

            ref_type = QLabel(reference.reference_type or "Custom")
            ref_type.setObjectName("ReferenceTypeBadge")

            title_row.addWidget(title, 1)
            title_row.addWidget(ref_type)

            url_label = QLabel(reference.url or "No URL provided.")
            url_label.setStyleSheet("font-size: 13px; color: #2563eb;")
            url_label.setWordWrap(True)

            note_label = QLabel(reference.note or "No additional note.")
            note_label.setStyleSheet("font-size: 13px; color: #374151;")
            note_label.setWordWrap(True)

            button_row = QHBoxLayout()
            open_button = QPushButton("Open Link")
            open_button.setObjectName("SecondaryButton")
            open_button.setFixedWidth(100)
            open_button.clicked.connect(lambda _=False, url=reference.url: self.open_reference_url(url))

            edit_button = QPushButton("Edit")
            edit_button.setObjectName("SecondaryButton")
            edit_button.setFixedWidth(80)
            edit_button.clicked.connect(lambda _=False, rid=reference.id: self.start_edit_reference(rid))

            delete_button = QPushButton("Delete")
            delete_button.setObjectName("DangerButton")
            delete_button.setFixedWidth(80)
            delete_button.clicked.connect(lambda _=False, rid=reference.id: self.delete_reference(rid))

            button_row.addWidget(open_button)
            button_row.addWidget(edit_button)
            button_row.addWidget(delete_button)
            button_row.addStretch()

            layout.addLayout(title_row)
            layout.addWidget(url_label)
            layout.addWidget(note_label)
            layout.addLayout(button_row)
            card.setLayout(layout)
            self.references_container_layout.addWidget(card)

        self.references_container_layout.addStretch()

    def handle_reference_auto_fill(self) -> None:
        title = self.reference_title_input.text().strip()
        ref_type = self.reference_type_input.currentText().strip()
        url = self.reference_url_input.text().strip()

        new_title, new_type, new_url = smart_fill_reference_fields(ref_type, title, url)
        self.reference_title_input.setText(new_title)
        self.reference_type_input.setCurrentText(new_type)
        self.reference_url_input.setText(new_url)

    def handle_reference_normalize_url(self) -> None:
        self.reference_url_input.setText(normalize_url(self.reference_url_input.text().strip()))

    def handle_save_reference(self) -> None:
        if self.current_selected_finding_id is None:
            QMessageBox.warning(self, "No Finding Selected", "Please select a finding first.")
            return

        title = self.reference_title_input.text().strip()
        if not title:
            QMessageBox.warning(self, "Missing Reference Title", "Please enter a reference title.")
            return

        session = SessionLocal()
        try:
            reference = Reference(
                finding_id=self.current_selected_finding_id,
                title=title,
                reference_type=self.reference_type_input.currentText().strip(),
                url=self.reference_url_input.text().strip(),
                note=self.reference_note_input.toPlainText().strip(),
            )
            session.add(reference)
            session.commit()
        finally:
            session.close()

        self.reset_reference_form()
        self.load_references()

    def start_edit_reference(self, reference_id: int) -> None:
        session = SessionLocal()
        try:
            reference = session.get(Reference, reference_id)
            if reference is None:
                return

            self.current_edit_reference_id = reference.id
            self.reference_title_input.setText(reference.title)
            self.reference_type_input.setCurrentText(reference.reference_type or "Custom")
            self.reference_url_input.setText(reference.url or "")
            self.reference_note_input.setPlainText(reference.note or "")

            self.reference_form_title.setText("Edit Reference")
            self.reference_form_subtitle.setText("Update the selected reference for the chosen finding.")

            self.save_reference_button.hide()
            self.clear_reference_form_button.hide()
            self.update_reference_button.show()
            self.cancel_reference_edit_button.show()

            self.nav_list.setCurrentRow(3)
        finally:
            session.close()

    def handle_update_reference(self) -> None:
        if self.current_edit_reference_id is None:
            return

        session = SessionLocal()
        try:
            reference = session.get(Reference, self.current_edit_reference_id)
            if reference is None:
                return

            reference.title = self.reference_title_input.text().strip()
            reference.reference_type = self.reference_type_input.currentText().strip()
            reference.url = self.reference_url_input.text().strip()
            reference.note = self.reference_note_input.toPlainText().strip()
            session.commit()
        finally:
            session.close()

        self.reset_reference_form()
        self.load_references()

    def delete_reference(self, reference_id: int) -> None:
        confirm = QMessageBox.question(self, "Delete Reference", "Delete this reference?")
        if confirm != QMessageBox.StandardButton.Yes:
            return

        session = SessionLocal()
        try:
            session.query(Reference).filter(Reference.id == reference_id).delete(synchronize_session=False)
            session.commit()
        finally:
            session.close()

        self.reset_reference_form()
        self.load_references()

    def open_reference_url(self, url: str) -> None:
        value = (url or "").strip()
        if not value:
            QMessageBox.warning(self, "Missing URL", "No URL is available for this reference.")
            return
        webbrowser.open(normalize_url(value))

    def collect_export_data(self):
        if self.current_project_id is None:
            return None, [], {}

        session = SessionLocal()
        try:
            project = session.get(Project, self.current_project_id)
            if project is None:
                return None, [], {}

            findings = (
                session.query(Finding)
                .filter(Finding.project_id == project.id)
                .order_by(Finding.id.asc())
                .all()
            )

            references = (
                session.query(Reference)
                .join(Finding, Reference.finding_id == Finding.id)
                .filter(Finding.project_id == project.id)
                .order_by(Reference.id.asc())
                .all()
            )

            references_by_finding: dict[int, list[Reference]] = defaultdict(list)
            for reference in references:
                references_by_finding[reference.finding_id].append(reference)

            return project, findings, references_by_finding
        finally:
            session.close()

    def build_checklist_page(self) -> QWidget:
        page = QWidget()

        page_scroll = QScrollArea()
        page_scroll.setWidgetResizable(True)
        page_scroll.setFrameShape(QFrame.Shape.NoFrame)

        page_content = QWidget()
        outer_layout = QVBoxLayout()
        outer_layout.setContentsMargins(0, 0, 0, 0)
        outer_layout.setSpacing(16)

        page_title = QLabel("Methodology Checklist")
        page_title.setStyleSheet("font-size: 30px; font-weight: bold;")

        page_subtitle = QLabel(
            "Track assessment execution and validate report readiness before final export."
        )
        page_subtitle.setStyleSheet("font-size: 14px; color: #4b5563;")

        self.checklist_project_card = QFrame()
        self.checklist_project_card.setObjectName("Card")
        checklist_project_layout = QVBoxLayout()
        checklist_project_layout.setContentsMargins(24, 20, 24, 20)
        checklist_project_layout.setSpacing(6)

        self.checklist_project_name_label = QLabel("No active project")
        self.checklist_project_name_label.setStyleSheet("font-size: 18px; font-weight: bold; color: #111827;")
        self.checklist_project_meta_label = QLabel("Select a current project to use the methodology checklist.")
        self.checklist_project_meta_label.setStyleSheet("font-size: 13px; color: #6b7280;")
        self.checklist_project_meta_label.setWordWrap(True)

        checklist_project_layout.addWidget(self.checklist_project_name_label)
        checklist_project_layout.addWidget(self.checklist_project_meta_label)
        self.checklist_project_card.setLayout(checklist_project_layout)

        actions_card = QFrame()
        actions_card.setObjectName("Card")
        actions_layout = QHBoxLayout()
        actions_layout.setContentsMargins(24, 18, 24, 18)
        actions_layout.setSpacing(10)

        mark_all_done_button = QPushButton("Mark All Done")
        mark_all_done_button.clicked.connect(self.mark_all_checklist_items_done)

        mark_all_not_required_button = QPushButton("Mark All Not Required")
        mark_all_not_required_button.setObjectName("SecondaryButton")
        mark_all_not_required_button.clicked.connect(self.mark_all_checklist_items_not_required)

        reset_checklist_button = QPushButton("Reset Checklist")
        reset_checklist_button.setObjectName("SecondaryButton")
        reset_checklist_button.clicked.connect(self.reset_checklist_items)

        jump_pending_button = QPushButton("Jump to Pending Items")
        jump_pending_button.setObjectName("SecondaryButton")
        jump_pending_button.clicked.connect(self.jump_to_pending_items)

        actions_layout.addWidget(mark_all_done_button)
        actions_layout.addWidget(mark_all_not_required_button)
        actions_layout.addWidget(reset_checklist_button)
        actions_layout.addWidget(jump_pending_button)
        actions_layout.addStretch()
        actions_card.setLayout(actions_layout)

        self.checklist_progress_label = QLabel("Progress: 0 / 0 completed")
        self.checklist_progress_label.setStyleSheet("font-size: 14px; font-weight: bold; color: #334155;")

        self.checklist_sections_container = QWidget()
        self.checklist_sections_layout = QVBoxLayout()
        self.checklist_sections_layout.setContentsMargins(0, 0, 0, 0)
        self.checklist_sections_layout.setSpacing(16)
        self.checklist_sections_container.setLayout(self.checklist_sections_layout)

        self.checklist_combo_by_key.clear()
        self.checklist_section_cards.clear()

        for section_name, items in CHECKLIST_SECTIONS:
            card = QFrame()
            card.setObjectName("Card")
            self.checklist_section_cards[section_name] = card

            card_layout = QVBoxLayout()
            card_layout.setContentsMargins(24, 20, 24, 20)
            card_layout.setSpacing(10)

            title = QLabel(section_name)
            title.setStyleSheet("font-size: 18px; font-weight: bold; color: #111827;")
            card_layout.addWidget(title)

            for item_key, item_label in items:
                row_frame = QFrame()
                row_frame.setObjectName("ChecklistRow")

                row_layout = QHBoxLayout()
                row_layout.setContentsMargins(12, 10, 12, 10)
                row_layout.setSpacing(12)

                label = QLabel(item_label)
                label.setWordWrap(True)
                label.setStyleSheet("font-size: 13px; color: #374151;")

                combo = QComboBox()
                combo.addItems(["To Do", "Done", "Not Required"])
                combo.setFixedWidth(160)
                combo.currentTextChanged.connect(
                    lambda value, key=item_key, section=section_name: self.handle_checklist_item_changed(section, key, value)
                )

                self.checklist_combo_by_key[item_key] = combo
                self.checklist_row_by_key[item_key] = row_frame
                self.checklist_label_by_key[item_key] = label
                self.checklist_section_by_key[item_key] = section_name

                row_layout.addWidget(label, 1)
                row_layout.addWidget(combo, 0, Qt.AlignmentFlag.AlignRight)

                row_frame.setLayout(row_layout)
                card_layout.addWidget(row_frame)

            card.setLayout(card_layout)
            self.checklist_sections_layout.addWidget(card)

        validation_card = QFrame()
        validation_card.setObjectName("Card")
        validation_layout = QVBoxLayout()
        validation_layout.setContentsMargins(24, 20, 24, 20)
        validation_layout.setSpacing(8)

        validation_title = QLabel("Checklist Validation")
        validation_title.setStyleSheet("font-size: 20px; font-weight: bold;")

        self.checklist_todo_label = QLabel("To Do remaining: 0")
        self.checklist_done_label = QLabel("Done: 0")
        self.checklist_not_required_label = QLabel("Not Required: 0")

        status_row = QHBoxLayout()
        status_row.setSpacing(18)

        self.checklist_status_label = QLabel("Checklist incomplete")
        self.checklist_status_label.setStyleSheet(
        "font-size: 20px; font-weight: bold; color: #b91c1c;"
        )

        self.checklist_pending_button = QPushButton("Pending checklist items")
        self.checklist_pending_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.checklist_pending_button.setFixedHeight(34)
        self.checklist_pending_button.clicked.connect(self.jump_to_pending_items)

        status_row.addWidget(self.checklist_status_label)
        status_row.addWidget(self.checklist_pending_button)
        status_row.addStretch()
        
        validation_layout.addWidget(validation_title)
        validation_layout.addWidget(self.checklist_todo_label)
        validation_layout.addWidget(self.checklist_done_label)
        validation_layout.addWidget(self.checklist_not_required_label)
        validation_layout.addSpacing(6)
        validation_layout.addLayout(status_row)
        validation_card.setLayout(validation_layout)

        outer_layout.addWidget(page_title)
        outer_layout.addWidget(page_subtitle)
        outer_layout.addWidget(self.checklist_project_card)
        outer_layout.addWidget(actions_card)
        outer_layout.addWidget(self.checklist_progress_label)
        outer_layout.addWidget(self.checklist_sections_container)
        outer_layout.addWidget(validation_card)
        outer_layout.addStretch()

        page_content.setLayout(outer_layout)
        page_scroll.setWidget(page_content)

        wrapper_layout = QVBoxLayout()
        wrapper_layout.setContentsMargins(0, 0, 0, 0)
        wrapper_layout.addWidget(page_scroll)
        page.setLayout(wrapper_layout)

        return page

    def update_checklist_context(self) -> None:
        session = SessionLocal()
        try:
            project = session.get(Project, self.current_project_id) if self.current_project_id else None
            if project is None:
                self.checklist_project_name_label.setText("No active project")
                self.checklist_project_meta_label.setText("Select a current project to use the methodology checklist.")
                self.load_checklist_state({})
                return

            findings_count = session.query(Finding).filter(Finding.project_id == project.id).count()
            self.checklist_project_name_label.setText(project.project_name)
            self.checklist_project_meta_label.setText(
                f"Client: {project.client_name or 'Not specified'}   •   "
                f"Target: {project.target_name or 'Not specified'}   •   "
                f"Findings: {findings_count}"
            )

            rows = (
                session.query(ChecklistItem)
                .filter(ChecklistItem.project_id == project.id)
                .all()
            )
            state_map = {row.item_key: row.status for row in rows}
            self.load_checklist_state(state_map)
        finally:
            session.close()

    def load_checklist_state(self, state_map: dict[str, str]) -> None:
        for key, combo in self.checklist_combo_by_key.items():
            combo.blockSignals(True)
            combo.setCurrentText(state_map.get(key, "To Do"))
            combo.blockSignals(False)
        self.refresh_checklist_summary()

    def handle_checklist_item_changed(self, section_name: str, item_key: str, value: str) -> None:
        if self.current_project_id is None:
            self.refresh_checklist_summary()
            return

        session = SessionLocal()
        try:
            row = (
                session.query(ChecklistItem)
                .filter(
                    ChecklistItem.project_id == self.current_project_id,
                    ChecklistItem.item_key == item_key,
                )
                .first()
            )

            if row is None:
                row = ChecklistItem(
                    project_id=self.current_project_id,
                    section_name=section_name,
                    item_key=item_key,
                    status=value,
                )
                session.add(row)
            else:
                row.status = value
                row.section_name = section_name

            session.commit()
        finally:
            session.close()

        self.refresh_checklist_summary()
        if self.pending_navigation_active:
             QTimer.singleShot(50, self._advance_to_next_pending_item)

    def refresh_checklist_summary(self) -> None:
        statuses = [combo.currentText() for combo in self.checklist_combo_by_key.values()]
        todo = sum(1 for status in statuses if status == "To Do")
        done = sum(1 for status in statuses if status == "Done")
        not_required = sum(1 for status in statuses if status == "Not Required")
        required_total = done + todo

        self.checklist_todo_label.setText(f"To Do remaining: {todo}")
        self.checklist_done_label.setText(f"Done: {done}")
        self.checklist_not_required_label.setText(f"Not Required: {not_required}")
        self.checklist_progress_label.setText(f"Progress: {done} / {required_total} completed")

        if todo == 0:
            self.checklist_status_label.setText("Checklist complete")
            self.checklist_status_label.setStyleSheet(
                "font-size: 20px; font-weight: bold; color: #166534;"
            )
            self.checklist_pending_button.setText("Ready for final export")
            self.checklist_pending_button.setStyleSheet(
                "color: #166534; background-color: #dcfce7; border: 1px solid #bbf7d0; border-radius: 10px; padding: 6px 14px; font-size: 14px; font-weight: 700;"
            )
            self.pending_navigation_active = False
            self.current_pending_key = None
        else:
            self.checklist_status_label.setText("Checklist incomplete")
            self.checklist_status_label.setStyleSheet(
                "font-size: 20px; font-weight: bold; color: #b91c1c;"
            )
            self.checklist_pending_button.setText("Pending checklist items")
            self.checklist_pending_button.setStyleSheet(
                "color: #b91c1c; background-color: #fee2e2; border: 1px solid #fecaca; border-radius: 10px; padding: 6px 14px; font-size: 14px; font-weight: 700;"
            )

        self._update_pending_highlights()

    def mark_all_checklist_items_done(self) -> None:
        for _, items in CHECKLIST_SECTIONS:
            for item_key, _ in items:
                combo = self.checklist_combo_by_key[item_key]
                combo.setCurrentText("Done")
        self.refresh_checklist_summary()

    def mark_all_checklist_items_not_required(self) -> None:
        for _, items in CHECKLIST_SECTIONS:
            for item_key, _ in items:
                combo = self.checklist_combo_by_key[item_key]
                combo.setCurrentText("Not Required")
        self.refresh_checklist_summary()

    def reset_checklist_items(self) -> None:
        for _, items in CHECKLIST_SECTIONS:
            for item_key, _ in items:
                combo = self.checklist_combo_by_key[item_key]
                combo.setCurrentText("To Do")
        self.refresh_checklist_summary()

    def jump_to_pending_items(self) -> None:
        pending_keys = self._pending_checklist_keys()

        if not pending_keys:
            QMessageBox.information(self, "Checklist Complete", "There are no pending checklist items.")
            self.pending_navigation_active = False
            self.current_pending_key = None
            self._update_pending_highlights()
            return

        self.pending_navigation_active = True
        self._update_pending_highlights()
        self._focus_checklist_item(pending_keys[0])

    def _pending_checklist_keys(self) -> list[str]:
        return [
            key
            for key, combo in self.checklist_combo_by_key.items()
            if combo.currentText() == "To Do"
        ]

    def _update_pending_highlights(self) -> None:
        pending_keys = set(self._pending_checklist_keys())

        for key, row in self.checklist_row_by_key.items():
            is_pending = key in pending_keys and self.pending_navigation_active
            row.setProperty("pending", is_pending)
            row.style().unpolish(row)
            row.style().polish(row)

            label = self.checklist_label_by_key.get(key)
            if label is not None:
                if is_pending:
                    label.setStyleSheet(
                        "font-size: 13px; font-weight: 700; color: #991b1b;"
                    )
                else:
                    label.setStyleSheet(
                        "font-size: 13px; color: #374151;"
                    )

    def _focus_checklist_item(self, item_key: str) -> None:
        combo = self.checklist_combo_by_key.get(item_key)
        if combo is None:
            return

        self.current_pending_key = item_key
        combo.setFocus()

    def _advance_to_next_pending_item(self) -> None:
        pending_keys = self._pending_checklist_keys()

        if not pending_keys:
            self.pending_navigation_active = False
            self.current_pending_key = None
            self._update_pending_highlights()
            return

        if self.current_pending_key in pending_keys:
            current_index = pending_keys.index(self.current_pending_key)
            next_index = min(current_index, len(pending_keys) - 1)
            self._focus_checklist_item(pending_keys[next_index])
        else:
            self._focus_checklist_item(pending_keys[0])

    def build_export_page(self) -> QWidget:
        page = QWidget()

        outer_layout = QVBoxLayout()
        outer_layout.setContentsMargins(0, 0, 0, 0)
        outer_layout.setSpacing(16)

        page_title = QLabel("Export")
        page_title.setStyleSheet("font-size: 30px; font-weight: bold;")

        page_subtitle = QLabel(
            "Export polished project reports and keep full control over the final filename and destination."
        )
        page_subtitle.setStyleSheet("font-size: 14px; color: #4b5563;")

        self.export_project_card = QFrame()
        self.export_project_card.setObjectName("Card")
 
        export_project_layout = QVBoxLayout()
        export_project_layout.setContentsMargins(24, 20, 24, 20)
        export_project_layout.setSpacing(6)
 
        self.export_project_name_label = QLabel("No active project")
        self.export_project_name_label.setStyleSheet(
            "font-size: 18px; font-weight: bold; color: #111827;"
        )

        self.export_project_meta_label = QLabel(
            "Go to Projects and select a current project first."
        )
        self.export_project_meta_label.setStyleSheet("font-size: 13px; color: #6b7280;")
        self.export_project_meta_label.setWordWrap(True)

        self.export_directory_label = QLabel("Last used export folder: not selected yet")
        self.export_directory_label.setStyleSheet("font-size: 13px; color: #4b5563;")

        choose_dir_button = QPushButton("Choose Export Folder")
        choose_dir_button.setObjectName("SecondaryButton")
        choose_dir_button.setFixedWidth(170)
        choose_dir_button.clicked.connect(self.choose_export_directory)

        export_buttons_row = QHBoxLayout()
        export_buttons_row.setSpacing(10)

        export_md_button = QPushButton("Export Markdown")
        export_md_button.setFixedWidth(160)
        export_md_button.clicked.connect(self.handle_export_markdown)

        export_txt_button = QPushButton("Export TXT")
        export_txt_button.setFixedWidth(120)
        export_txt_button.clicked.connect(self.handle_export_text)

        export_docx_button = QPushButton("Export DOCX")
        export_docx_button.setFixedWidth(130)
        export_docx_button.clicked.connect(self.handle_export_docx)

        export_pdf_button = QPushButton("Export PDF")
        export_pdf_button.setFixedWidth(120)
        export_pdf_button.clicked.connect(self.handle_export_pdf)

        open_folder_button = QPushButton("Open Export Folder")
        open_folder_button.setObjectName("SecondaryButton")
        open_folder_button.setFixedWidth(160)
        open_folder_button.clicked.connect(self.open_export_folder)

        export_buttons_row.addWidget(export_md_button)
        export_buttons_row.addWidget(export_txt_button)
        export_buttons_row.addWidget(export_docx_button)
        export_buttons_row.addWidget(export_pdf_button)
        export_buttons_row.addWidget(open_folder_button)
        export_buttons_row.addStretch()

        export_project_layout.addWidget(self.export_project_name_label)
        export_project_layout.addWidget(self.export_project_meta_label)
        export_project_layout.addWidget(self.export_directory_label)
        export_project_layout.addWidget(choose_dir_button, alignment=Qt.AlignmentFlag.AlignLeft)
        export_project_layout.addLayout(export_buttons_row)

        self.export_project_card.setLayout(export_project_layout)

        outer_layout.addWidget(page_title)
        outer_layout.addWidget(page_subtitle)
        outer_layout.addWidget(self.export_project_card)
        outer_layout.addStretch()

        page.setLayout(outer_layout)
        return page
    
    def choose_export_directory(self) -> None:
        start_dir = str(self.last_export_directory) if self.last_export_directory else ""
        selected_dir = QFileDialog.getExistingDirectory(self, "Choose Export Folder", start_dir)
        if not selected_dir:
            return
        self.last_export_directory = Path(selected_dir)
        self.update_export_context()

    def update_export_context(self) -> None:
        session = SessionLocal()
        try:
            project = None
            if self.current_project_id is not None:
                project = session.get(Project, self.current_project_id)

            if project is None:
                self.export_project_name_label.setText("No active project")
                self.export_project_meta_label.setText(
                    "Go to Projects and select a current project first."
                )
            else:
                findings_count = session.query(Finding).filter(Finding.project_id == project.id).count()
                self.export_project_name_label.setText(project.project_name)
                self.export_project_meta_label.setText(
                    f"Client: {project.client_name or 'Not specified'}   •   "
                    f"Target: {project.target_name or 'Not specified'}   •   "
                    f"Type: {project.engagement_type or 'Not specified'}   •   "
                    f"Risk: {project.risk_rating or 'Not Rated'}   •   "
                    f"Findings: {findings_count}"
                )

            if self.last_export_directory is None:
                self.export_directory_label.setText("Last used export folder: not selected yet")
            else:
                self.export_directory_label.setText(f"Last used export folder: {self.last_export_directory}")
        finally:
            session.close()

    def ask_export_file_path(self, suggested_name: str, filter_text: str) -> Path | None:
        initial_dir = str(self.last_export_directory) if self.last_export_directory else ""
        suggested_path = str(Path(initial_dir) / suggested_name) if initial_dir else suggested_name

        selected_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Report As",
            suggested_path,
            filter_text,
        )
        if not selected_path:
            return None

        final_path = Path(selected_path)
        final_path.parent.mkdir(parents=True, exist_ok=True)
        self.last_export_directory = final_path.parent
        self.update_export_context()
        return final_path

    def _export_with_save_as(self, extension: str, exporter_func) -> None:
        project, findings, references_by_finding = self.collect_export_data()
        if project is None:
            QMessageBox.warning(self, "No Active Project", "Please select a current project before exporting.")
            return

        suggested_name = f"{project.project_name}.{extension}".replace("/", "-").replace("\\", "-")
        filter_text = f"{extension.upper()} Files (*.{extension})"
        final_path = self.ask_export_file_path(suggested_name, filter_text)
        if final_path is None:
            return

        try:
            with TemporaryDirectory() as temp_dir:
                temp_output = exporter_func(Path(temp_dir), project, findings, references_by_finding)
                shutil.copy2(temp_output, final_path)

            QMessageBox.information(self, "Export Complete", f"Report created successfully:\n{final_path}")
        except Exception as error:
            QMessageBox.critical(self, "Export Error", f"An error occurred during export:\n{error}")

    def handle_export_markdown(self) -> None:
        self._export_with_save_as("md", export_markdown_report)

    def handle_export_text(self) -> None:
        self._export_with_save_as("txt", export_text_report)

    def handle_export_docx(self) -> None:
        self._export_with_save_as("docx", export_docx_report)

    def handle_export_pdf(self) -> None:
        self._export_with_save_as("pdf", export_pdf_report)

    def open_export_folder(self) -> None:
        if self.last_export_directory is None:
            QMessageBox.warning(self, "No Export Folder", "Choose an export folder first.")
            return
        webbrowser.open(self.last_export_directory.as_uri())