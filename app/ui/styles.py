APP_STYLE = """
QMainWindow {
    background-color: #f5f7fb;
}

QLabel {
    color: #1f2937;
}

QFrame#Card {
    background-color: #ffffff;
    border: 1px solid #d9e2f1;
    border-radius: 16px;
}

QFrame#ArchivePanel {
    background-color: #f8fafc;
    border: 1px solid #dbe5f2;
    border-radius: 12px;
}

QFrame#ProjectItemCard {
    background-color: #f8fafc;
    border: 1px solid #dbe5f2;
    border-radius: 12px;
}

QFrame#ProjectItemCard[selected="true"] {
    background-color: #eaf2ff;
    border: 2px solid #2563eb;
    border-radius: 12px;
}

QLabel#MonthGroupHeader {
    font-size: 14px;
    font-weight: 700;
    color: #334155;
    padding-top: 10px;
    padding-bottom: 2px;
}

QLabel#SeverityBadgeCritical {
    background-color: #fee2e2;
    color: #991b1b;
    border: 1px solid #fecaca;
    border-radius: 8px;
    padding: 4px 10px;
    font-weight: 700;
}

QLabel#SeverityBadgeHigh {
    background-color: #ffedd5;
    color: #9a3412;
    border: 1px solid #fed7aa;
    border-radius: 8px;
    padding: 4px 10px;
    font-weight: 700;
}

QLabel#SeverityBadgeMedium {
    background-color: #fef3c7;
    color: #92400e;
    border: 1px solid #fde68a;
    border-radius: 8px;
    padding: 4px 10px;
    font-weight: 700;
}

QLabel#SeverityBadgeLow {
    background-color: #dcfce7;
    color: #166534;
    border: 1px solid #bbf7d0;
    border-radius: 8px;
    padding: 4px 10px;
    font-weight: 700;
}

QLabel#SeverityBadgeInfo {
    background-color: #dbeafe;
    color: #1d4ed8;
    border: 1px solid #bfdbfe;
    border-radius: 8px;
    padding: 4px 10px;
    font-weight: 700;
}

QLabel#ReferenceTypeBadge {
    background-color: #eef2ff;
    color: #4338ca;
    border: 1px solid #c7d2fe;
    border-radius: 8px;
    padding: 4px 10px;
    font-weight: 700;
}

QPushButton {
    background-color: #2563eb;
    color: white;
    border: none;
    border-radius: 10px;
    padding: 10px 16px;
    font-weight: 600;
    min-height: 20px;
}

QPushButton:hover {
    background-color: #1d4ed8;
}

QPushButton#SecondaryButton {
    background-color: #e5e7eb;
    color: #111827;
    border: 1px solid #d1d5db;
}

QPushButton#SecondaryButton:hover {
    background-color: #dbe1e8;
}

QPushButton#DangerButton {
    background-color: #dc2626;
    color: white;
    border: none;
}

QPushButton#DangerButton:hover {
    background-color: #b91c1c;
}

QListWidget {
    background-color: transparent;
    border: none;
    outline: none;
    font-size: 14px;
    color: #1f2937;
    padding: 4px;
}

QListWidget#AttachmentList,
QListWidget#FindingSelectionList {
    background-color: #ffffff;
    border: 1px solid #cbd5e1;
    border-radius: 10px;
    padding: 8px;
}

QListWidget::item {
    padding: 10px 12px;
    border-radius: 8px;
    margin-bottom: 4px;
}

QListWidget::item:selected {
    background-color: #dbeafe;
    color: #1d4ed8;
}

QListWidget::item:hover {
    background-color: #eef4ff;
}

QLineEdit, QTextEdit {
    background-color: #ffffff;
    border: 1px solid #cbd5e1;
    border-radius: 10px;
    padding: 10px;
    font-size: 14px;
    color: #1f2937;
}

QComboBox {
    background-color: #f3f4f6;
    border: 1px solid #cbd5e1;
    border-radius: 10px;
    padding: 10px;
    font-size: 14px;
    color: #1f2937;
}

QComboBox:focus {
    border: 1px solid #2563eb;
}

QComboBox::drop-down {
    border: none;
    width: 28px;
    background: transparent;
}

QComboBox QAbstractItemView {
    background-color: #f3f4f6;
    color: #1f2937;
    border: 1px solid #cbd5e1;
    selection-background-color: #dbeafe;
    selection-color: #1d4ed8;
    outline: none;
    padding: 4px;
}

QLineEdit:focus, QTextEdit:focus {
    border: 1px solid #2563eb;
}

QScrollArea {
    background: transparent;
    border: none;
}

QScrollArea > QWidget > QWidget {
    background: transparent;
}

QScrollBar:vertical {
    background: transparent;
    width: 10px;
    margin: 0px;
}

QScrollBar::handle:vertical {
    background: #94a3b8;
    border-radius: 5px;
    min-height: 20px;
}

QScrollBar::add-line:vertical,
QScrollBar::sub-line:vertical {
    height: 0px;
}

QScrollBar:horizontal {
    background: transparent;
    height: 10px;
    margin: 0px;
}

QScrollBar::handle:horizontal {
    background: #94a3b8;
    border-radius: 5px;
    min-width: 20px;
}

QScrollBar::add-line:horizontal,
QScrollBar::sub-line:horizontal {
    width: 0px;
}

QMessageBox {
    background-color: #111827;
}

QMessageBox QLabel {
    color: #f9fafb;
}

QMessageBox QPushButton {
    min-width: 80px;
}

QLabel#ChecklistStatusGood {
    color: #166534;
    font-weight: 700;
}

QLabel#ChecklistStatusBad {
    color: #b91c1c;
    font-weight: 700;
}

QFrame#ChecklistRow {
    background-color: transparent;
    border: 1px solid transparent;
    border-radius: 10px;
}

QFrame#ChecklistRow[pending="true"] {
    background-color: #fef2f2;
    border: 1px solid #fecaca;
    border-radius: 10px;
}

QPushButton#LinkButton {
    background: transparent;
    border: none;
    color: #b91c1c;
    padding: 0px;
    margin: 0px;
    font-size: 16px;
    font-weight: 700;
    text-align: left;
}

QPushButton#LinkButton:hover {
    background: transparent;
    color: #991b1b;
    text-decoration: underline;
}
"""