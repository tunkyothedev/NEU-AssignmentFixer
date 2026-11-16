import os
import time
import win32com.client
from typing import Optional, List, Dict

class WordController:
    """
    Quản lý tự động hóa Microsoft Word qua COM interface (pywin32).
    Hỗ trợ so sánh tài liệu (Track Changes), chèn nhận xét (Comments) và chuyển đổi PDF -> DOCX.
    """
    def __init__(self, visible: bool = False):
        self.visible = visible
        self.word = None

    def __enter__(self):
        # Sử dụng DispatchEx để luôn khởi tạo tiến trình Word độc lập, không bị ảnh hưởng bởi tiến trình Word cũ đang treo
        self.word = win32com.client.DispatchEx("Word.Application")
        self.word.Visible = self.visible
        self.word.DisplayAlerts = 0  # wdAlertsNone
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.word:
            try:
                self.word.Quit(SaveChanges=0)  # wdDoNotSaveChanges
            except Exception:
                pass
            self.word = None

    def compare_documents(
        self,
        original_path: str,
        revised_path: str,
        output_path: str,
        author_name: str = "AssignmentFixer"
    ) -> bool:
        """
        So sánh file gốc và file đã sửa để sinh file có Track Changes nguyên bản của Word.
        """
        orig_doc = None
        rev_doc = None
        compared_doc = None
        try:
            abs_orig = os.path.abspath(original_path)
            abs_rev = os.path.abspath(revised_path)
            abs_out = os.path.abspath(output_path)

            orig_doc = self.word.Documents.Open(abs_orig, ReadOnly=True)
            rev_doc = self.word.Documents.Open(abs_rev, ReadOnly=True)

            # CompareDocuments(OriginalDocument, RevisedDocument, Destination, Granularity, ...)
            # wdCompareDestinationNew = 2, wdGranularityWordLevel = 1
            compared_doc = self.word.CompareDocuments(
                OriginalDocument=orig_doc,
                RevisedDocument=rev_doc,
                Destination=2,
                Granularity=1,
                CompareFormatting=True,
                CompareCaseChanges=True,
                CompareWhitespace=False,
                CompareTables=True,
                CompareHeaders=True,
                CompareFootnotes=True,
                CompareTextboxes=True,
                CompareFields=True,
                CompareComments=True,
                RevisedAuthor=author_name
            )

            compared_doc.SaveAs2(abs_out, FileFormat=12)  # wdFormatXMLDocument (.docx)
            return True
        finally:
            if compared_doc:
                try:
                    compared_doc.Close(SaveChanges=0)
                except Exception:
                    pass
            if rev_doc:
                try:
                    rev_doc.Close(SaveChanges=0)
                except Exception:
                    pass
            if orig_doc:
                try:
                    orig_doc.Close(SaveChanges=0)
                except Exception:
                    pass

    def add_comments(self, doc_path: str, comments: List[Dict[str, str]]) -> bool:
        """
        Chèn các nhận xét (Comments) vào tài liệu Word theo các từ khóa / đoạn văn bản.
        comments: list of {'target_text': str, 'comment_text': str}
        """
        doc = None
        try:
            abs_path = os.path.abspath(doc_path)
            doc = self.word.Documents.Open(abs_path)
            
            for item in comments:
                target = item.get("target_text", "").strip()
                text = item.get("comment_text", "").strip()
                if not target or not text:
                    continue

                find_obj = doc.Content.Find
                find_obj.ClearFormatting()
                find_obj.Text = target[:200]  # Giới hạn độ dài tìm kiếm
                find_obj.Forward = True
                find_obj.Wrap = 0  # wdFindStop

                if find_obj.Execute():
                    doc.Comments.Add(Range=find_obj.Parent, Text=text)
            
            doc.Save()
            return True
        finally:
            if doc:
                try:
                    doc.Close(SaveChanges=1)
                except Exception:
                    pass

    def convert_pdf_to_docx(self, pdf_path: str, docx_path: str) -> bool:
        """
        Dùng bộ chuyển đổi Word Reflow của Microsoft Word để chuyển đổi PDF sang DOCX nguyên bản.
        """
        doc = None
        try:
            abs_pdf = os.path.abspath(pdf_path)
            abs_docx = os.path.abspath(docx_path)
            
            doc = self.word.Documents.Open(abs_pdf, ConfirmConversions=False)
            doc.SaveAs2(abs_docx, FileFormat=12)
            return True
        finally:
            if doc:
                try:
                    doc.Close(SaveChanges=0)
                except Exception:
                    pass
