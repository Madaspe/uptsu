from docxtpl import DocxTemplate, RichText
from docxcompose.composer import Composer
from docx import Document

from os import makedirs


def mkdir(result):
    try:
        makedirs('/'.join(result.split("/")[:-1]))
    except:
        pass


def unite_docx(files, result):
    merged_document = Document()

    mkdir(result)

    for index, file in enumerate(files):
        sub_doc = Document(file)

        # Don't add a page break if you've reached the last file.
        if index < len(files)-1:
           sub_doc.add_page_break()

        for element in sub_doc.element.body:
            merged_document.element.body.append(element)

    merged_document.save(result)


def generate_doc(data, source, result):
    tpl = DocxTemplate(source)

    tpl.render(data)

    mkdir(result)

    tpl.save(result)


if __name__ == '__main__':
    generate_doc({
        "fio": "Алферов Никита Дмитриевич",
        "group": "БИУ-22-01"
    }, "templates/dest.docx")
