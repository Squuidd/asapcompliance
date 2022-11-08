from venv import create
from docxtpl import DocxTemplate
from docxtpl.subdoc import Subdoc, SubdocComposer
from docx import Document
from zipfile import ZipFile

from lxml import etree  # TODO for debug
import re
import io

import types

import pythoncom
import win32com.client
import tempfile

import os

# import field_updater
# UNO = field_updater.UNO(
#     "127.0.0.1", "60000"
# )

# import win32com.client

# import inspect, os


def findPath(file_name):
    script_dir = os.path.dirname(__file__)  # absolute dir the script is in
    rel_path = f"Safety Programs/{file_name}"
    abs_file_path = os.path.join(script_dir, rel_path)
    return abs_file_path

# def update_toc(docx_file):
#     word = win32com.client.DispatchEx("Word.Application")
#     doc = word.Documents.Open(docx_file)
#     doc.TablesOfContents(1).Update()
#     doc.Close(SaveChanges=True)
#     word.Quit()

class Tempdoc():
    def __init__(self, data: bytes, filetype: str = "docx", word=None):
        if not word:
            word = self.word_instance()
            self.word_given = False # need to close word if it wasnt given
        else:
            self.word_given = True

        self.path = self.make_temp(data, filetype)
        self.doc = word.Documents.Open(self.path)

        self.word = word

        # This is msword monkey business
        self.format = {
            "docx": 16,
            "pdf": 17
        }

    @staticmethod
    def word_instance():
        pythoncom.CoInitialize()
        word = win32com.client.DispatchEx("Word.Application")
        return word

    def make_temp(self, data, filetype):
        Counter = 0
        TEMPPATH = tempfile.gettempdir()
        while True:
            path = f"{TEMPPATH}\\{str(Counter)}.{filetype}"
            try:
                with open(path, 'wb') as f:
                    f.write(data)
                break
            except:
                Counter += 1
        return path

    def save_as(self, _format="pdf"):
        # Good idea to make an empty file and overwrite later
        path = self.make_temp(b"", _format)

        self.doc.SaveAs(
            path,
            FileFormat=self.format[_format]
        )
        return self.read(path)

    def save(self):
        self.doc.Close(SaveChanges=True)
        return self.read(self.path)

    def read(self, path):
        with open(path) as f:
            data = f.buffer.read()
        return data

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        try:
            self.doc.Close()
            os.remove(self.path)
        except:
            pass

        if not self.word_given:
            self.word.Quit()

def DocumentBytes(doc):
    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    return file_stream.getvalue()


class DummyDoc(Subdoc):
    def __init__(self, tpl, xml):
        super().__init__(tpl)  # tpl is main document, it is normally passed to subdoc class
        self.xml = xml

    def _get_xml(self):  # Spoof the xml with our own
        return self.xml


def create_manual(
        file,
        safety_documents,
        company_name
):
    main_document = DocxTemplate(file)
    # Create subdoc to insert into main document
    main_document.init_docx()  # Not sure what the point of this is but lib normally does it before init Subdoc.

    xml = ""
    compose = SubdocComposer(main_document)
    for doc in safety_documents:
        sd = Document(doc)
        compose.attach_parts(sd)

        # Remove any sections because it breaks shit
        if sd.element.body.sectPr is not None:
            sd.element.body.remove(sd.element.body.sectPr)

        # add the bodies of every subdoc to our xml
        xml += re.sub(r'</?w:body[^>]*>', '', etree.tostring(
            sd.element.body, encoding='unicode', pretty_print=False))

    subdoc = DummyDoc(main_document, xml)

    ctx = {
        "safety": subdoc,
        "company_name": company_name
    }

    main_document.render(ctx)  # Render

    with Tempdoc(DocumentBytes(main_document)) as td:
        td.doc.TablesOfContents(1).Update()
        b = td.save()
    return b

def create_program(
        files: list,  # TODO Should maybe be bytes
        company_name: str
):
    docs = []
    wd = Tempdoc.word_instance() # slow to reopen and close word repeatedly
    for file in files:
        main_document = DocxTemplate(file)

        ctx = {
            "company_name": company_name
        }

        main_document.render(ctx)

        # docx
        data = DocumentBytes(main_document)
        docs.append(
            [
                os.path.basename(file),
                data
            ],
        )

        # pdf
        with Tempdoc(data) as td:
            data_pdf = td.save_as("pdf")
        docs.append(
            [
                os.path.basename(file)[:-4] + "pdf",
                data_pdf
            ],
        )

            
    return zip_files(docs).getvalue()

def zip_files(files):
    mem_zip = io.BytesIO()
    with ZipFile(mem_zip, mode="w") as zf:
        for file in files:
            zf.writestr(file[0], file[1])
    zf.close()
    return mem_zip

# create_manual(findPath("safety_manual.docx"), [findPath("aerial lifts.docx"), findPath("cranes.docx"), findPath("cadmium.docx")], "Test Name LLC.")
# update_toc('Output/new_safety_manual.docx')

# create_program(findPath("aerial lifts.docx"), "Test Name.")
