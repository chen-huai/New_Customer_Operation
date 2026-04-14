DEFAULT_THEME = "light_blue"

THEME_PALETTES = {
    "light_blue": {
        "window": "#f4f8ff",
        "surface": "#ffffff",
        "surface_alt": "#eef4ff",
        "border": "#c8d8f0",
        "text": "#1f2937",
        "muted": "#5b6472",
        "accent": "#2f80ed",
        "accent_hover": "#2469c4",
        "accent_pressed": "#1c569f",
        "disabled": "#9aa5b1",
    },
    "dark_blue": {
        "window": "#152232",
        "surface": "#1c2d42",
        "surface_alt": "#22364d",
        "border": "#3b5673",
        "text": "#e6eefb",
        "muted": "#a8b7ca",
        "accent": "#58a6ff",
        "accent_hover": "#4394ef",
        "accent_pressed": "#2f78cf",
        "disabled": "#748396",
    },
    "light_green": {
        "window": "#f3fcf6",
        "surface": "#ffffff",
        "surface_alt": "#eaf8ef",
        "border": "#c6dfcf",
        "text": "#213229",
        "muted": "#617067",
        "accent": "#2f9e6b",
        "accent_hover": "#27865a",
        "accent_pressed": "#1e6a47",
        "disabled": "#9cad9f",
    },
}


class ThemeManager:
    def __init__(self, app):
        self.app = app
        self.current_theme = DEFAULT_THEME
        self.themes = list(THEME_PALETTES.keys())

    def is_available(self):
        return bool(self.themes)

    def set_theme(self, theme):
        if theme not in THEME_PALETTES:
            return self.set_default_theme()

        self.current_theme = theme
        self.app.setStyleSheet(self._build_stylesheet(THEME_PALETTES[theme]))
        return True

    def set_default_theme(self):
        return self.set_theme(DEFAULT_THEME)

    def toggle_theme(self):
        current_index = self.themes.index(self.current_theme)
        next_index = (current_index + 1) % len(self.themes)
        return self.set_theme(self.themes[next_index])

    def get_available_themes(self):
        return self.themes

    def _build_stylesheet(self, palette):
        return f"""
            QWidget {{
                background-color: {palette["window"]};
                color: {palette["text"]};
            }}
            QMainWindow, QMenuBar, QMenu, QStatusBar, QTabWidget::pane, QGroupBox {{
                background-color: {palette["surface"]};
                color: {palette["text"]};
                border: 1px solid {palette["border"]};
            }}
            QGroupBox {{
                margin-top: 12px;
                padding-top: 14px;
                border-radius: 10px;
                font-weight: 600;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 4px;
            }}
            QMenuBar {{
                padding: 4px;
            }}
            QMenuBar::item {{
                background: transparent;
                padding: 6px 10px;
                border-radius: 6px;
            }}
            QMenuBar::item:selected,
            QMenu::item:selected,
            QTabBar::tab:selected {{
                background-color: {palette["surface_alt"]};
                color: {palette["text"]};
            }}
            QMenu::item {{
                padding: 6px 24px 6px 12px;
            }}
            QTabBar::tab {{
                background-color: {palette["surface_alt"]};
                color: {palette["muted"]};
                border: 1px solid {palette["border"]};
                border-bottom: none;
                border-top-left-radius: 8px;
                border-top-right-radius: 8px;
                padding: 7px 14px;
                margin-right: 4px;
            }}
            QPushButton {{
                background-color: {palette["accent"]};
                color: #ffffff;
                border: none;
                border-radius: 8px;
                padding: 8px 14px;
                font-weight: 600;
            }}
            QPushButton:hover {{
                background-color: {palette["accent_hover"]};
            }}
            QPushButton:pressed {{
                background-color: {palette["accent_pressed"]};
            }}
            QPushButton:disabled,
            QPushButton[enabled="false"] {{
                background-color: {palette["border"]};
                color: {palette["disabled"]};
            }}
            QLineEdit, QTextBrowser, QComboBox, QSpinBox, QDoubleSpinBox {{
                background-color: {palette["surface_alt"]};
                color: {palette["text"]};
                border: 1px solid {palette["border"]};
                border-radius: 8px;
                padding: 6px 8px;
                selection-background-color: {palette["accent"]};
                selection-color: #ffffff;
            }}
            QComboBox QAbstractItemView {{
                background-color: {palette["surface"]};
                color: {palette["text"]};
                border: 1px solid {palette["border"]};
                selection-background-color: {palette["accent"]};
                selection-color: #ffffff;
            }}
            QLabel {{
                color: {palette["text"]};
            }}
        """
