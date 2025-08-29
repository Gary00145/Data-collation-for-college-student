"""UI样式定义文件"""

# 颜色定义
COLORS = {
    'primary': '#3498db',        # 主蓝色
    'success': '#2ecc71',        # 成功绿色
    'warning': '#f39c12',        # 警告橙色
    'danger': '#e74c3c',         # 危险红色
    'dark': '#333333',           # 深灰色文字
    'light': '#f5f7fa',          # 浅灰色背景
    'white': '#ffffff',          # 白色
    'border': '#dddddd',         # 边框颜色
}

# 样式表
STYLESHEET = f"""
* {{
    font-family: "Microsoft YaHei", "SimHei", sans-serif;
    font-size: 10pt;
}}

QMainWindow {{
    background-color: {COLORS['light']};
}}

QMenuBar {{
    background-color: {COLORS['white']};
    border-bottom: 1px solid {COLORS['border']};
    font-size: 10pt;
}}

QMenuBar::item {{
    background: transparent;
    padding: 4px 8px;
}}

QMenuBar::item:selected {{
    background: {COLORS['primary']};
    color: {COLORS['white']};
}}

QMenuBar::item:pressed {{
    background: {COLORS['primary']};
    color: {COLORS['white']};
}}

QMenu {{
    background-color: {COLORS['white']};
    border: 1px solid {COLORS['border']};
    font-size: 10pt;
}}

QMenu::item {{
    padding: 6px 20px;
}}

QMenu::item:selected {{
    background-color: {COLORS['primary']};
    color: {COLORS['white']};
}}

QPushButton {{
    background-color: {COLORS['primary']};
    color: {COLORS['white']};
    border: none;
    padding: 8px 16px;
    border-radius: 4px;
    font-weight: bold;
    font-size: 10pt;
}}

QPushButton:hover {{
    background-color: #2980b9;
}}

QPushButton:pressed {{
    background-color: #21618c;
}}

QPushButton#success {{
    background-color: {COLORS['success']};
}}

QPushButton#success:hover {{
    background-color: #27ae60;
}}

QPushButton#warning {{
    background-color: {COLORS['warning']};
}}

QPushButton#warning:hover {{
    background-color: #e67e22;
}}

QPushButton#danger {{
    background-color: {COLORS['danger']};
}}

QPushButton#danger:hover {{
    background-color: #c0392b;
}}

QListWidget, QTreeWidget, QTextEdit {{
    background-color: {COLORS['white']};
    border: 1px solid {COLORS['border']};
    border-radius: 4px;
    font-size: 10pt;
}}

QListWidget::item:selected, QTreeWidget::item:selected {{
    background-color: {COLORS['primary']};
    color: {COLORS['white']};
}}

QLabel {{
    color: {COLORS['dark']};
    font-weight: bold;
    font-size: 10pt;
}}

QStatusBar {{
    background-color: {COLORS['white']};
    border-top: 1px solid {COLORS['border']};
    font-size: 9pt;
}}

QProgressBar {{
    border: 1px solid {COLORS['border']};
    border-radius: 4px;
    text-align: center;
    font-size: 9pt;
}}

QProgressBar::chunk {{
    background-color: {COLORS['primary']};
    width: 20px;
}}

QComboBox {{
    background-color: {COLORS['white']};
    border: 1px solid {COLORS['border']};
    border-radius: 4px;
    padding: 4px;
    font-size: 10pt;
}}

QComboBox:hover {{
    border-color: {COLORS['primary']};
}}

QDialog {{
    background-color: {COLORS['light']};
}}

QLineEdit, QTextEdit {{
    border: 1px solid {COLORS['border']};
    border-radius: 4px;
    padding: 6px;
    font-size: 10pt;
}}

QLineEdit:focus, QTextEdit:focus {{
    border-color: {COLORS['primary']};
}}

/* 标题标签使用稍大的字体 */
QLabel#fileLabel, QLabel#treeLabel, QLabel#titleLabel, QLabel#contentLabel {{
    font-size: 11pt;
    margin-bottom: 5px;
}}
"""