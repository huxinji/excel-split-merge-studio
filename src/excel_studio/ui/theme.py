from __future__ import annotations


def product_stylesheet() -> str:
    """High-legibility light material system inspired by Apple Liquid Glass."""

    return """
    * {
        font-family: "Microsoft YaHei UI", "PingFang SC", "Noto Sans SC", sans-serif;
        font-size: 13px;
        color: #172033;
    }
    QMainWindow, QWidget#AppRoot {
        background: qlineargradient(
            x1: 0, y1: 0, x2: 1, y2: 1,
            stop: 0 #F9FBFF,
            stop: 0.48 #F2F6FC,
            stop: 1 #EAF2FB
        );
    }

    /* Floating glass navigation layer. */
    QFrame#ProductHeader {
        background: qlineargradient(
            x1: 0, y1: 0, x2: 0, y2: 1,
            stop: 0 rgba(255, 255, 255, 248),
            stop: 1 rgba(242, 247, 253, 238)
        );
        border: none;
        border-bottom: 1px solid #CDD8E6;
    }
    QLabel#ProductTitle {
        color: #101828;
        font-size: 22px;
        font-weight: 700;
    }
    QLabel#ProductSubtitle {
        color: #45566F;
        font-size: 12px;
        font-weight: 500;
    }
    QLabel#VersionBadge {
        color: #174EA6;
        background: rgba(236, 245, 255, 235);
        border: 1px solid #B9D6F7;
        border-radius: 12px;
        padding: 4px 10px;
        font-size: 11px;
        font-weight: 600;
    }
    QFrame#WorkspaceSwitch {
        background: rgba(226, 234, 244, 205);
        border: 1px solid #C3CFDE;
        border-radius: 17px;
    }
    QPushButton#WorkspaceButton {
        color: #34445C;
        background: transparent;
        border: none;
        border-radius: 14px;
        padding: 7px 17px;
        min-height: 28px;
        font-weight: 650;
    }
    QPushButton#WorkspaceButton:hover {
        background: rgba(255, 255, 255, 205);
        color: #174EA6;
    }
    QPushButton#WorkspaceButton:checked {
        background: #FFFFFF;
        color: #075FBD;
        border: 1px solid #B9D2EF;
    }
    QComboBox#LanguageCombo {
        color: #24334A;
        background: rgba(255, 255, 255, 225);
        border: 1px solid #B9C7D9;
        border-radius: 15px;
        padding: 7px 30px 7px 12px;
        min-width: 112px;
        min-height: 28px;
        font-weight: 600;
    }
    QComboBox#LanguageCombo:hover {
        background: #FFFFFF;
        border-color: #7EA9DA;
    }
    QComboBox#LanguageCombo QAbstractItemView {
        color: #172033;
        background: #FFFFFF;
        border: 1px solid #C7D3E1;
        selection-background-color: #DCEEFF;
        selection-color: #102A4C;
    }

    QScrollArea, QScrollArea > QWidget > QWidget {
        background: transparent;
        border: none;
    }
    QSplitter::handle {
        background: transparent;
        width: 8px;
    }

    /* Stable content surfaces: opaque enough for dependable readability. */
    QFrame#StepCard {
        background: rgba(255, 255, 255, 248);
        border: 1px solid #CED9E7;
        border-radius: 16px;
    }
    QFrame#StepCard[accent="blue"] { border-top: 3px solid #4092EE; }
    QFrame#StepCard[accent="cyan"] { border-top: 3px solid #25A6C7; }
    QFrame#StepCard[accent="violet"] { border-top: 3px solid #8A6BE8; }
    QFrame#StepCard[accent="amber"] { border-top: 3px solid #D99A32; }
    QFrame#StepCard[accent="emerald"] { border-top: 3px solid #35A878; }
    QFrame#StepCard[accent="rose"] { border-top: 3px solid #D96888; }
    QLabel#StepNumber {
        color: #FFFFFF;
        background: qlineargradient(
            x1: 0, y1: 0, x2: 1, y2: 1,
            stop: 0 #2B86E8,
            stop: 1 #0068D9
        );
        border: 1px solid #005FC7;
        border-radius: 13px;
        min-width: 26px;
        min-height: 26px;
        max-width: 26px;
        max-height: 26px;
        font-weight: 700;
    }
    QLabel#CardTitle {
        color: #14213A;
        font-size: 15px;
        font-weight: 700;
    }
    QLabel#CardDescription, QLabel#Muted {
        color: #52627A;
        font-size: 12px;
        font-weight: 450;
    }
    QLabel#PrimaryRuleTitle {
        color: #075FBD;
        font-size: 16px;
        font-weight: 700;
    }
    QFrame#PrimaryRulePanel {
        background: #F3F8FE;
        border: 1px solid #AFC9E8;
        border-radius: 13px;
    }

    /* Segmented and capsule controls form the glass control layer. */
    QFrame#SegmentedControl {
        background: rgba(226, 234, 244, 210);
        border: 1px solid #C5D1DF;
        border-radius: 14px;
    }
    QPushButton#SegmentButton {
        background: transparent;
        color: #3B4B62;
        border: none;
        border-radius: 11px;
        padding: 7px 13px;
        min-height: 28px;
        font-weight: 650;
    }
    QPushButton#SegmentButton:hover {
        background: rgba(255, 255, 255, 210);
        color: #075FBD;
    }
    QPushButton#SegmentButton:checked {
        background: qlineargradient(
            x1: 0, y1: 0, x2: 0, y2: 1,
            stop: 0 #1677E8,
            stop: 1 #0068D9
        );
        color: #FFFFFF;
        border: 1px solid #005FC7;
    }
    QPushButton {
        min-height: 28px;
        font-weight: 600;
    }
    QPushButton#PrimaryButton {
        color: #FFFFFF;
        background: qlineargradient(
            x1: 0, y1: 0, x2: 0, y2: 1,
            stop: 0 #1677E8,
            stop: 1 #0068D9
        );
        border: 1px solid #005FC7;
        border-radius: 13px;
        padding: 7px 17px;
        min-height: 30px;
        font-size: 13px;
        font-weight: 700;
    }
    QPushButton#PrimaryButton:hover {
        background: #075FBD;
        border-color: #0755A7;
    }
    QPushButton#PrimaryButton:pressed {
        background: #054E9C;
    }
    QPushButton#PrimaryButton:disabled {
        color: #68768A;
        background: #E3E9F1;
        border-color: #C6D0DD;
    }
    QPushButton#SecondaryButton {
        color: #174EA6;
        background: rgba(240, 247, 255, 235);
        border: 1px solid #ABC8E9;
        border-radius: 12px;
        padding: 6px 12px;
        min-height: 28px;
        font-weight: 650;
    }
    QPushButton#SecondaryButton:hover {
        color: #0A438D;
        background: #FFFFFF;
        border-color: #6FA4DD;
    }
    QPushButton#DangerButton {
        color: #B42318;
        background: #FFF5F4;
        border: 1px solid #E9B8B2;
        border-radius: 12px;
        padding: 6px 12px;
        min-height: 28px;
        font-weight: 650;
    }
    QPushButton#DangerButton:hover {
        background: #FFE8E5;
        border-color: #D98D84;
    }
    QPushButton#SuccessButton {
        color: #FFFFFF;
        background: #087A55;
        border: 1px solid #066B4A;
        border-radius: 12px;
        padding: 6px 13px;
        min-height: 28px;
        font-weight: 700;
    }
    QPushButton#SuccessButton:hover {
        background: #066B4A;
    }

    QLineEdit, QComboBox, QSpinBox {
        color: #172033;
        background: rgba(255, 255, 255, 245);
        border: 1px solid #B8C6D7;
        border-radius: 11px;
        padding: 5px 8px;
        min-height: 24px;
        selection-background-color: #B9DBFF;
        selection-color: #10233F;
    }
    QLineEdit:hover, QComboBox:hover, QSpinBox:hover {
        border-color: #86A8CE;
        background: #FFFFFF;
    }
    QLineEdit:focus, QComboBox:focus, QSpinBox:focus {
        background: #FFFFFF;
        border: 2px solid #287FD9;
    }
    QLineEdit:disabled, QComboBox:disabled, QSpinBox:disabled {
        color: #667085;
        background: #E9EEF5;
        border-color: #CBD4E0;
    }
    QComboBox::drop-down {
        border: none;
        width: 25px;
    }
    QCheckBox {
        spacing: 7px;
        color: #26364E;
        min-height: 24px;
    }
    QCheckBox:disabled {
        color: #667085;
    }

    QTableWidget, QListWidget, QPlainTextEdit {
        color: #172033;
        background: #FFFFFF;
        alternate-background-color: #F5F8FC;
        border: 1px solid #C3D0DF;
        border-radius: 11px;
        gridline-color: #DDE5EF;
        selection-background-color: #D6E9FF;
        selection-color: #10233F;
    }
    QTableWidget:disabled, QListWidget:disabled, QPlainTextEdit:disabled {
        color: #667085;
        background: #EEF2F7;
    }
    QHeaderView::section {
        color: #243A59;
        background: #E8F0F9;
        border: none;
        border-right: 1px solid #CBD7E5;
        border-bottom: 1px solid #B9C8D9;
        padding: 7px;
        font-weight: 700;
    }

    QFrame#InspectionPanel {
        background: rgba(255, 255, 255, 232);
        border: 1px solid #C9D5E3;
        border-radius: 16px;
    }
    QLabel#MetricValue {
        color: #14213A;
        font-size: 17px;
        font-weight: 700;
    }
    QLabel#MetricLabel {
        color: #52627A;
        font-size: 11px;
        font-weight: 500;
    }

    /* Fixed glass action layer, deliberately more opaque for text contrast. */
    QFrame#FooterBar {
        background: qlineargradient(
            x1: 0, y1: 0, x2: 0, y2: 1,
            stop: 0 rgba(255, 255, 255, 248),
            stop: 1 rgba(238, 244, 251, 244)
        );
        border-top: 1px solid #BCCADB;
    }
    QFrame#FooterBar QLabel {
        color: #34445C;
        font-weight: 550;
    }
    QFrame#FooterBar QLineEdit {
        color: #172033;
        background: #FFFFFF;
        border: 1px solid #AABBD0;
    }
    QProgressBar {
        color: #24334A;
        background: #DCE5F0;
        border: 1px solid #CAD5E2;
        border-radius: 4px;
        text-align: center;
        min-height: 8px;
        max-height: 8px;
    }
    QProgressBar::chunk {
        background: #1677E8;
        border-radius: 3px;
    }

    QScrollBar:vertical {
        background: transparent;
        width: 11px;
        margin: 2px;
    }
    QScrollBar::handle:vertical {
        background: #AEBED1;
        border-radius: 4px;
        min-height: 28px;
    }
    QScrollBar::handle:vertical:hover {
        background: #879BB3;
    }
    QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
        height: 0;
    }
    QScrollBar:horizontal {
        background: transparent;
        height: 11px;
        margin: 2px;
    }
    QScrollBar::handle:horizontal {
        background: #AEBED1;
        border-radius: 4px;
        min-width: 28px;
    }
    QScrollBar::handle:horizontal:hover {
        background: #879BB3;
    }
    QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
        width: 0;
    }
    QToolTip {
        color: #FFFFFF;
        background: #26364E;
        border: 1px solid #0F1D30;
        border-radius: 7px;
        padding: 6px;
    }
    """
