"""
Estilos CSS/QSS para la aplicación
"""

MAIN_STYLESHEET = """
    QWidget {
        background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #F0F4F8, stop:1 #E5E9F0);
        font-family: 'Inter', 'Segoe UI', sans-serif;
    }
    QPushButton {
        background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #667EEA, stop:1 #564FEE);
        color: white; border: none; padding: 12px 24px; border-radius: 10px;
        font-weight: 600; font-size: 13px; min-height: 40px;
    }
    QPushButton:hover {
        background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #7C8FF5, stop:1 #6B64F8);
    }
    QPushButton:disabled {
        background: #D1D5DB; color: #9CA3AF;
    }
    
    /* TABS */
    QTabWidget::pane {
        border: none;
        background: white;
        border-radius: 0 0 12px 12px;
    }
    QTabBar {
        background: #F9FAFB;
    }
    QTabBar::tab {
        background: #E5E7EB;
        color: #374151;
        padding: 12px 24px;
        margin-right: 4px;
        margin-top: 4px;
        border: none;
        border-radius: 8px 8px 0 0;
        font-size: 13px;
        font-weight: 600;
        min-width: 100px;
    }
    QTabBar::tab:selected {
        background: white;
        color: #667EEA;
        border-bottom: 3px solid #667EEA;
    }
    QTabBar::tab:hover:!selected {
        background: #D1D5DB;
        color: #1F2937;
    }
    
    /* TABLES */
    QTableView {
        background: white;
        border: none;
        gridline-color: #F3F4F6;
        selection-background-color: #EEF2FF;
        selection-color: #1E293B;
        font-size: 12px;
        alternate-background-color: #F9FAFB;
    }
    QTableView::item {
        padding: 8px;
        border-bottom: 1px solid #F3F4F6;
        color: #1F2937;
    }
    QTableView::item:selected {
        background: #EEF2FF;
        color: #1E293B;
    }
    QHeaderView::section {
        background: #F9FAFB;
        color: #374151;
        padding: 12px 8px;
        border: none;
        border-bottom: 2px solid #E5E7EB;
        font-weight: 700;
        font-size: 11px;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    
    /* CHECKBOXES */
    QCheckBox {
        spacing: 10px;
        color: #374151;
        font-size: 14px;
        font-weight: 600;
        padding: 8px;
    }
    QCheckBox::indicator {
        width: 22px;
        height: 22px;
        border-radius: 6px;
        border: 2px solid #D1D5DB;
        background: white;
    }
    QCheckBox::indicator:hover {
        border-color: #667EEA;
    }
    QCheckBox::indicator:checked {
        background: #667EEA;
        border-color: #667EEA;
    }
    
    /* TEXT EDIT */
    QTextEdit {
        background: white;
        border: none;
        padding: 12px;
        font-family: 'Cascadia Code', 'Consolas', monospace;
        font-size: 12px;
        color: #374151;
        line-height: 1.5;
    }
    
    /* PROGRESS BAR */
    QProgressBar {
        border: none;
        border-radius: 3px;
        background: #E5E7EB;
        height: 6px;
    }
    QProgressBar::chunk {
        background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #667EEA, stop:1 #764FEE);
        border-radius: 3px;
    }
"""

CARD_STYLESHEET = """
    ModernCard {
        background: white; 
        border-radius: 16px; 
        border: 1px solid #E5E7EB;
    }
"""