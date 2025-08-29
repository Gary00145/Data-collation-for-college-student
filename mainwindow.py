import os
import subprocess
import sys
import tempfile
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (QApplication, QMainWindow, QFileDialog, QTreeWidget,
                             QTreeWidgetItem, QListWidget, QPushButton, QHBoxLayout,
                             QVBoxLayout, QWidget, QLabel, QSplitter, QComboBox,
                             QAction, QMessageBox, QMenu, QProgressDialog, QDialog,
                             QTextEdit, QLineEdit, QDialogButtonBox)
from document_processor import DocumentProcessor
from previewwindow import PreviewWindow
from ai_processor import AIProcessor
from styles import STYLESHEET, COLORS


class MainWindow(QMainWindow):
    """主应用窗口"""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("大学生资料整理平台")
        # 使用相对尺寸设置窗口大小
        self.resize(1200, 800)
        screen = QApplication.primaryScreen().availableGeometry()
        self.setGeometry(
            (screen.width() - 1200) // 2,
            (screen.height() - 800) // 2,
            1200,
            800
        )
        
        # 应用样式表
        self.setStyleSheet(STYLESHEET)

        # 数据结构
        self.documents = []
        self.knowledge_tree = []

        # 创建UI
        self.init_ui()

        # 预览窗口
        self.preview_window = PreviewWindow()

        # 添加上下文菜单
        self.setup_context_menu()

    def setup_context_menu(self):
        """添加上下文菜单"""
        # 文件列表的右键菜单
        self.file_list.setContextMenuPolicy(Qt.CustomContextMenu)
        self.file_list.customContextMenuRequested.connect(self.show_file_context_menu)

        # 知识树的右键菜单
        self.tree_widget.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tree_widget.customContextMenuRequested.connect(self.show_tree_context_menu)

    def show_file_context_menu(self, position):
        """显示文件列表的右键菜单"""
        menu = QMenu()

        # 删除选中文件
        delete_action = QAction("删除选中文件", self)
        delete_action.triggered.connect(self.delete_selected_files)
        menu.addAction(delete_action)

        # 清空所有文件
        clear_action = QAction("清空所有文件", self)
        clear_action.triggered.connect(self.clear_all_files)
        menu.addAction(clear_action)

        menu.exec_(self.file_list.mapToGlobal(position))

    def show_tree_context_menu(self, position):
        """显示知识树的右键菜单"""
        menu = QMenu()

        # 删除选中节点
        delete_action = QAction("删除选中节点", self)
        delete_action.triggered.connect(self.delete_selected_node)
        menu.addAction(delete_action)

        # 编辑节点内容
        edit_action = QAction("编辑节点内容", self)
        edit_action.triggered.connect(self.edit_selected_node)
        menu.addAction(edit_action)

        menu.exec_(self.tree_widget.mapToGlobal(position))

    def delete_selected_files(self):
        """删除选中的文件"""
        selected_items = self.file_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "警告", "请先选择要删除的文件")
            return

        # 确认删除
        reply = QMessageBox.question(
            self,
            "确认删除",
            f"确定要删除选中的 {len(selected_items)} 个文件吗？",
            QMessageBox.Yes | QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            # 从数据结构中移除
            for item in selected_items:
                file_name = item.text()
                # 找到对应的文档并从数据结构中移除
                self.documents = [doc for doc in self.documents if os.path.basename(doc['filepath']) != file_name]
                # 从列表中移除
                self.file_list.takeItem(self.file_list.row(item))

            self.update_status(f"已删除 {len(selected_items)} 个文件")
            # 清除知识树
            self.knowledge_tree = []
            self.tree_widget.clear()

    def clear_all_files(self):
        """清空所有文件"""
        if self.file_list.count() == 0:
            return

        # 确认清除
        reply = QMessageBox.question(
            self,
            "确认清除",
            "确定要清空所有文件吗？",
            QMessageBox.Yes | QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            self.file_list.clear()
            self.documents = []
            self.knowledge_tree = []
            self.tree_widget.clear()
            self.update_status("已清空所有文件")

    def delete_selected_node(self):
        """删除选中的知识树节点"""
        selected_items = self.tree_widget.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "警告", "请先选择要删除的节点")
            return

        # 确认删除
        reply = QMessageBox.question(
            self,
            "确认删除",
            f"确定要删除选中的 {len(selected_items)} 个节点吗？",
            QMessageBox.Yes | QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            # 从知识树中移除
            for item in selected_items:
                node_index = self.tree_widget.indexOfTopLevelItem(item)
                if 0 <= node_index < len(self.knowledge_tree):
                    del self.knowledge_tree[node_index]
                self.tree_widget.takeTopLevelItem(self.tree_widget.indexOfTopLevelItem(item))

            self.update_status(f"已删除 {len(selected_items)} 个节点")

    def edit_selected_node(self):
        """编辑选中的节点"""
        selected_items = self.tree_widget.selectedItems()
        if not selected_items or len(selected_items) > 1:
            QMessageBox.warning(self, "警告", "请选择单个节点进行编辑")
            return

        item = selected_items[0]
        node_index = self.tree_widget.indexOfTopLevelItem(item)
        if 0 <= node_index < len(self.knowledge_tree):
            node_data = self.knowledge_tree[node_index]
            # 打开编辑对话框
            dialog = NodeEditDialog(node_data, self)
            if dialog.exec_() == QDialog.Accepted:
                # 更新节点数据
                updated_data = dialog.get_data()
                self.knowledge_tree[node_index] = updated_data
                # 更新UI显示
                item.setText(0, f"{node_index + 1}. {updated_data['title']}")
                item.setData(0, Qt.UserRole, updated_data)
                self.update_status("节点已更新")

    def init_ui(self):
        """初始化用户界面"""
        # 主布局
        main_widget = QWidget()
        main_layout = QHBoxLayout()
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(15, 15, 15, 15)

        # 左侧面板：文件操作
        left_panel = QVBoxLayout()
        left_panel.setSpacing(10)

        # 文件列表标题和列表
        file_label = QLabel("已上传文件:")
        file_label.setObjectName("fileLabel")
        self.file_list = QListWidget()
        self.file_list.setMinimumHeight(200)

        # 操作按钮
        btn_upload = QPushButton("上传文件")
        btn_upload.clicked.connect(self.upload_files)
        btn_upload.setMinimumHeight(35)

        btn_generate = QPushButton("生成知识结构")
        btn_generate.clicked.connect(self.generate_knowledge_tree)
        btn_generate.setMinimumHeight(35)

        # AI处理按钮
        btn_ai_process = QPushButton("AI二次总结")
        btn_ai_process.clicked.connect(self.process_with_ai)
        btn_ai_process.setMinimumHeight(35)

        # 导出选项
        export_layout = QHBoxLayout()
        export_label = QLabel("导出格式:")
        export_label.setObjectName("exportLabel")
        self.export_combo = QComboBox()
        self.export_combo.addItems(["导出为Word", "导出为PDF"])
        self.export_combo.setMinimumHeight(30)
        btn_export = QPushButton("导出")
        btn_export.clicked.connect(self.export_document)
        btn_export.setMinimumHeight(35)
        export_layout.addWidget(export_label)
        export_layout.addWidget(self.export_combo)
        export_layout.addWidget(btn_export)

        left_panel.addWidget(file_label)
        left_panel.addWidget(self.file_list)
        left_panel.addWidget(btn_upload)
        left_panel.addWidget(btn_generate)
        left_panel.addWidget(btn_ai_process)
        left_panel.addLayout(export_layout)

        # 右侧面板：知识树和预览
        right_panel = QVBoxLayout()
        right_panel.setSpacing(10)

        # 知识结构标题和树
        tree_label = QLabel("知识结构:")
        tree_label.setObjectName("treeLabel")
        self.tree_widget = QTreeWidget()
        self.tree_widget.setHeaderLabel("知识结构")
        self.tree_widget.itemClicked.connect(self.show_preview)
        self.tree_widget.setMinimumHeight(300)

        right_panel.addWidget(tree_label)
        right_panel.addWidget(self.tree_widget)

        # 组合布局
        splitter = QSplitter(Qt.Horizontal)
        splitter.setHandleWidth(8)
        
        left_widget = QWidget()
        left_widget.setLayout(left_panel)
        right_widget = QWidget()
        right_widget.setLayout(right_panel)

        splitter.addWidget(left_widget)
        splitter.addWidget(right_widget)
        splitter.setSizes([300, 900])

        # 主窗口布局
        main_layout.addWidget(splitter)
        main_widget.setLayout(main_layout)
        self.setCentralWidget(main_widget)

    def upload_files(self):
        """文件上传功能"""
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "选择学习资料",
            "",
            "文档文件 (*.pdf *.docx *.pptx)"
        )

        if not files:
            return

        for file_path in files:
            # 添加到文件列表
            self.file_list.addItem(os.path.basename(file_path))
            # 解析文档结构
            self.process_file(file_path)

    def process_file(self, file_path):
        """处理单个文件并存储数据结构"""
        processor = DocumentProcessor()
        ext = os.path.splitext(file_path)[1].lower()

        try:
            if ext == '.pdf':
                content = processor.extract_pdf_content(file_path)
            elif ext == '.docx':
                content = processor.extract_docx_content(file_path)
            elif ext == '.pptx':
                content = processor.extract_pptx_content(file_path)
            else:
                return

            content['filepath'] = file_path
            self.documents.append(content)
            self.update_status(f"已加载: {os.path.basename(file_path)}")
        except Exception as e:
            self.update_status(f"错误: {str(e)}")

    def generate_knowledge_tree(self):
        """生成知识结构树"""
        if not self.documents:
            self.update_status("请先上传文件")
            return

        processor = DocumentProcessor()
        self.knowledge_tree = processor.generate_knowledge_tree(self.documents)

        # 更新树状视图
        self.tree_widget.clear()
        for i, node in enumerate(self.knowledge_tree):
            item = QTreeWidgetItem([f"{i + 1}. {node['title']}"])
            item.setData(0, Qt.UserRole, node)
            self.tree_widget.addTopLevelItem(item)

        self.update_status("知识结构已生成")

    def process_with_ai(self):
        """使用AI进行二次总结"""
        if not self.knowledge_tree:
            self.update_status("请先生成知识结构")
            return

        # 创建进度对话框
        progress = QProgressDialog("正在使用AI进行二次总结...", "取消", 0, 3, self)
        progress.setWindowModality(Qt.WindowModal)
        progress.setWindowTitle("AI处理中")
        progress.show()

        try:
            # 初始化AI处理器
            progress.setLabelText("初始化AI处理器...")
            progress.setValue(1)
            ai_processor = AIProcessor()

            # 准备内容
            progress.setLabelText("准备内容...")
            progress.setValue(2)
            content = ai_processor.prepare_content_for_ai(self.knowledge_tree)

            # 发送到AI处理
            progress.setLabelText("发送到AI进行处理...")
            progress.setValue(3)
            ai_response = ai_processor.send_to_ai(content)

            # 解析AI响应
            self.knowledge_tree = ai_processor.parse_ai_response(ai_response)

            # 更新树状视图
            self.tree_widget.clear()
            for i, node in enumerate(self.knowledge_tree):
                item = QTreeWidgetItem([f"{i + 1}. {node['title']}"])
                item.setData(0, Qt.UserRole, node)
                self.tree_widget.addTopLevelItem(item)

            self.update_status("AI二次总结完成")
            progress.close()

        except Exception as e:
            progress.close()
            QMessageBox.critical(self, "AI处理失败", f"AI处理过程中出现错误: {str(e)}")
            self.update_status(f"AI处理失败: {str(e)}")

    def show_preview(self, item, column):
        """显示文件预览"""
        node_data = item.data(0, Qt.UserRole)
        filepath = node_data.get('filepath', '')

        if not filepath:
            return

        ext = os.path.splitext(filepath)[1].lower()

        self.preview_window.show()
        try:
            if ext == '.pdf':
                self.preview_window.show_pdf_preview(filepath)
            elif ext == '.docx':
                self.preview_window.show_docx_preview(filepath)
            elif ext == '.pptx':
                self.preview_window.show_pptx_preview(filepath)
        except Exception as e:
            self.update_status(f"预览失败: {str(e)}")

    def export_document(self):
        """导出整理后的文档（优化过滤器默认选中状态）"""
        if not self.knowledge_tree:
            self.update_status("请先生成知识结构")
            return

        processor = DocumentProcessor()
        format_choice = self.export_combo.currentText()  # 获取下拉框选择的格式（"导出为Word"或"导出为PDF"）

        # 根据选择的格式，设置对应的过滤器和默认文件名
        if "Word" in format_choice:
            default_filename = "知识结构总结.docx"
            file_filter = "Word文档 (*.docx);;PDF文件 (*.pdf)"
            selected_filter_index = 0  # 默认选中第1个过滤器（Word）
        else:  # PDF格式
            default_filename = "知识结构总结.pdf"
            file_filter = "PDF文件 (*.pdf);;Word文档 (*.docx)"  # 交换顺序，让PDF在前面
            selected_filter_index = 0  # 默认选中第1个过滤器（PDF）

        # 获取文件对话框的选中过滤器（通过split分割过滤器字符串，取对应索引的过滤器）
        filters = file_filter.split(';;')
        selected_filter = filters[selected_filter_index]

        # 打开文件保存对话框，指定默认过滤器
        output_file, _ = QFileDialog.getSaveFileName(
            self,
            "保存整理文档",
            default_filename,
            file_filter,
            selected_filter  # 关键：指定默认选中的过滤器
        )

        if not output_file:
            return  # 用户取消保存

        try:
            if "Word" in format_choice:
                # 处理Word导出（确保扩展名正确）
                if not output_file.lower().endswith('.docx'):
                    output_file += '.docx'
                result = processor.export_to_word(self.knowledge_tree, output_file)
            else:
                # 处理PDF导出（先生成临时Word，再转换）
                temp_word_path = os.path.join(tempfile.gettempdir(), f"temp_pdf_conv_{os.getpid()}.docx")
                processor.export_to_word(self.knowledge_tree, temp_word_path)

                # 确保PDF扩展名正确
                if not output_file.lower().endswith('.pdf'):
                    output_file += '.pdf'
                result = processor.export_to_pdf(temp_word_path, output_file)

                # 清理临时文件
                if os.path.exists(temp_word_path):
                    os.remove(temp_word_path)

            if result and os.path.exists(result):
                self.update_status(f"导出成功: {os.path.basename(result)}")
                # 自动打开文件
                if sys.platform == "win32":
                    os.startfile(result)
                else:
                    opener = "open" if sys.platform == "darwin" else "xdg-open"
                    subprocess.call([opener, result])
            else:
                self.update_status("导出失败，未生成文件")
        except Exception as e:
            self.update_status(f"导出失败: {str(e)}")

    def update_status(self, message):
        """更新状态栏"""
        self.statusBar().showMessage(message, 5000)


class NodeEditDialog(QDialog):
    """节点编辑对话框"""

    def __init__(self, node_data, parent=None):
        super().__init__(parent)
        self.node_data = node_data.copy()  # 复制数据以避免直接修改原始数据
        self.init_ui()

    def init_ui(self):
        """初始化UI"""
        self.setWindowTitle("编辑节点")
        self.setModal(True)
        self.resize(500, 400)
        # 应用样式
        self.setStyleSheet(STYLESHEET)

        layout = QVBoxLayout()
        layout.setSpacing(10)
        layout.setContentsMargins(20, 20, 20, 20)

        # 标题输入
        title_label = QLabel("标题:")
        title_label.setObjectName("titleLabel")
        self.title_edit = QLineEdit(self.node_data.get('title', ''))
        self.title_edit.setMinimumHeight(30)

        # 内容输入
        content_label = QLabel("内容:")
        content_label.setObjectName("contentLabel")
        self.content_edit = QTextEdit()
        content = self.node_data.get('content', '')
        if isinstance(content, list):
            content = '\n'.join(content)
        self.content_edit.setPlainText(content)
        self.content_edit.setMinimumHeight(200)

        # 按钮框
        button_box = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel
        )
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        ok_button = button_box.button(QDialogButtonBox.Ok)
        cancel_button = button_box.button(QDialogButtonBox.Cancel)
        ok_button.setMinimumHeight(35)
        cancel_button.setMinimumHeight(35)

        layout.addWidget(title_label)
        layout.addWidget(self.title_edit)
        layout.addWidget(content_label)
        layout.addWidget(self.content_edit)
        layout.addWidget(button_box)

        self.setLayout(layout)

    def get_data(self):
        """获取编辑后的数据"""
        return {
            'title': self.title_edit.text(),
            'content': self.content_edit.toPlainText(),
            'children': self.node_data.get('children', [])
        }


# This is the main program that runs the GUI
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())