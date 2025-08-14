import re
import fitz
import pdfplumber
from docx import Document
from pdfminer.layout import LAParams
from pptx import Presentation


class DocumentProcessor:
    """文档处理核心模块 - 完全移除 PyMuPDF 依赖"""

    @staticmethod
    def extract_pdf_content(filepath):
        """
        兼容最新 pdfplumber 版本的 PDF 解析方法
        """
        content = {"metadata": {}, "sections": []}

        # 优先使用 PyMuPDF (fitz) 解析器
        result = DocumentProcessor.extract_with_pymupdf(filepath)
        if result and result.get('sections') and result['sections'][0].get('content'):
            return result

        # 如果 PyMuPDF 失败，尝试 pdfplumber
        print("PyMuPDF 解析不完整，尝试 pdfplumber")
        return DocumentProcessor.extract_with_pdfplumber(filepath)

    @staticmethod
    def extract_with_pymupdf(filepath):
        """使用 PyMuPDF 作为主要解析器"""
        try:
            content = {"metadata": {}, "sections": []}
            current_section = None

            doc = fitz.open(filepath)
            content['metadata']['pages'] = len(doc)
            content['metadata']['author'] = doc.metadata.get('author', 'Unknown')

            for page_num in range(len(doc)):
                page = doc.load_page(page_num)
                try:
                    # PyMuPDF 通常能更好地处理颜色空间问题
                    text = page.get_text("text")

                    if text:
                        # 清理页码和特殊字符
                        cleaned_text = DocumentProcessor.clean_content(text)

                        # 处理内容结构
                        lines = cleaned_text.split('\n')
                        for line in lines:
                            stripped_line = line.strip()
                            if stripped_line:
                                # 如果还没有创建节，或者遇到可能的标题
                                if current_section is None or DocumentProcessor.is_heading(stripped_line):
                                    if current_section and current_section['content']:
                                        content['sections'].append(current_section)
                                    current_section = {
                                        'title': stripped_line,
                                        'content': []
                                    }
                                else:
                                    current_section['content'].append(stripped_line)
                except Exception as e:
                    print(f"第 {page_num + 1} 页 PyMuPDF 解析错误: {str(e)}")
                    if current_section is None:
                        current_section = {"title": "解析错误", "content": []}
                    current_section['content'].append(f"第 {page_num + 1} 页解析失败")

            # 添加最后一节
            if current_section and current_section['content']:
                content['sections'].append(current_section)

            return content

        except Exception as e:
            print(f"PyMuPDF 完全失败: {str(e)}")
            return None

    @staticmethod
    def extract_with_pdfplumber(filepath):
        """使用 pdfplumber 解析，兼容最新版本"""
        content = {"metadata": {}, "sections": []}
        current_section = None

        try:
            # 使用 pdfminer 的 LAParams
            laparams = LAParams()

            with pdfplumber.open(filepath, laparams=laparams) as pdf:
                content['metadata']['pages'] = len(pdf.pages)
                content['metadata']['author'] = pdf.metadata.get('Author', 'Unknown')

                for i, page in enumerate(pdf.pages):
                    try:
                        # 尝试标准提取
                        text = page.extract_text()
                    except Exception as e:
                        # 处理颜色空间错误
                        if "invalid color" in str(e).lower() or "gray" in str(e).lower():
                            print(f"第 {i + 1} 页遇到颜色空间错误，尝试简化提取")
                            try:
                                # 尝试简化提取方法
                                text = page.extract_text_simple()
                            except:
                                text = ""
                                print(f"第 {i + 1} 页简化提取失败")
                        else:
                            text = ""
                            print(f"第 {i + 1} 页解析错误: {str(e)}")

                    if text:
                        # 清理页码
                        cleaned_text = DocumentProcessor.clean_content(text)

                        # 处理内容结构
                        lines = cleaned_text.split('\n')
                        for line in lines:
                            stripped_line = line.strip()
                            if stripped_line:
                                if current_section is None or DocumentProcessor.is_heading(stripped_line):
                                    if current_section and current_section['content']:
                                        content['sections'].append(current_section)
                                    current_section = {
                                        'title': stripped_line,
                                        'content': []
                                    }
                                else:
                                    if current_section is not None:
                                        current_section['content'].append(stripped_line)

            # 添加最后一节
            if current_section and current_section['content']:
                content['sections'].append(current_section)

            return content

        except Exception as e:
            print(f"pdfplumber 解析失败: {str(e)}")
            # 尝试使用 PyMuPDF 作为后备
            return DocumentProcessor.extract_with_pymupdf(filepath)

    @staticmethod
    def clean_content(text):
        """清理内容：去除页码和其他不需要的元素"""
        if not text:
            return ""

        # 1. 去除单独一行的数字（页码）
        cleaned = re.sub(r'^\s*\d+\s*$', '', text, flags=re.MULTILINE)

        # 2. 去除类似 "P123" 的错误信息残留
        cleaned = re.sub(r'\bP\d+\b', '', cleaned)

        # 3. 去除常见的页码模式 (Page X, P.X, etc.)
        cleaned = re.sub(r'(?i)\b(?:page|p|pg|pag|pagina)\s*[\.:]*\s*\d+\b', '', cleaned)

        # 4. 去除多余空行
        cleaned = re.sub(r'\n\s*\n', '\n', cleaned)

        return cleaned.strip()

    @staticmethod
    def is_heading(text):
        """判断是否为标题的简单实现"""
        # 可以根据您的文档特点扩展此方法
        heading_keywords = ["目录", "章节", "节", "第", "摘要", "引言", "结论", "参考"]
        if len(text) < 50 and any(keyword in text for keyword in heading_keywords):
            return True
        return False


    # 其他extract方法保持不变
    @staticmethod
    def extract_docx_content(filepath):
        """提取Word文档内容"""
        doc = Document(filepath)
        content = {"metadata": {}, "sections": []}
        current_section = {"title": "引言", "content": []}

        for para in doc.paragraphs:
            if para.style.name.startswith('Heading'):
                # 新章节开始
                if current_section['content']:
                    content['sections'].append(current_section)
                current_section = {
                    'title': para.text.strip(),
                    'content': [para.text]
                }
            else:
                current_section['content'].append(para.text)

        if current_section['content']:
            content['sections'].append(current_section)
        return content

    @staticmethod
    def extract_pptx_content(filepath):
        """提取PPT内容"""
        prs = Presentation(filepath)
        content = {"metadata": {}, "sections": []}

        for i, slide in enumerate(prs.slides):
            slide_title = f"幻灯片 {i + 1}"
            slide_content = []

            # 提取幻灯片标题
            if slide.shapes.title:
                slide_title = slide.shapes.title.text

            # 提取内容
            for shape in slide.shapes:
                if hasattr(shape, "text") and shape != slide.shapes.title:
                    slide_content.append(shape.text)

            content['sections'].append({
                'title': slide_title,
                'content': '\n'.join(slide_content)
            })
        return content

    # generate_knowledge_tree 保持不变
    @staticmethod
    def generate_knowledge_tree(documents):
        knowledge_tree = []
        for doc in documents:
            for section in doc['sections']:
                node = {
                    'title': section['title'],
                    'content': '\n'.join(section['content']),
                    'children': []
                }
                knowledge_tree.append(node)
        return knowledge_tree

    # export_to_word 保持不变
    @staticmethod
    def export_to_word(knowledge_tree, output_path):
        doc = Document()

        for node in knowledge_tree:
            doc.add_heading(node['title'], level=1)
            if isinstance(node['content'], list):
                for item in node['content']:
                    doc.add_paragraph(item)
            else:
                doc.add_paragraph(node['content'])

        doc.save(output_path)
        return output_path

    # export_to_pdf 保持不变（使用reportlab）
    @staticmethod
    def export_to_pdf(word_path, pdf_path):
        try:
            import comtypes.client
            import os

            # 确保路径为绝对路径
            word_path = os.path.abspath(word_path)
            pdf_path = os.path.abspath(pdf_path)

            # 创建Word应用实例（后台运行，不显示窗口）
            word = comtypes.client.CreateObject('Word.Application')
            word.Visible = False

            # 打开Word文档
            doc = word.Documents.Open(word_path)

            # 保存为PDF（FileFormat=17是Word中PDF的固定格式代码）
            doc.SaveAs2(pdf_path, FileFormat=17)

            # 关闭文档和Word进程
            doc.Close()
            word.Quit()

            return pdf_path if os.path.exists(pdf_path) else None
        except Exception as e:
            print(f"PDF导出失败: {str(e)}")
            return None
