from docxtpl import DocxTemplate, RichText
from docxcompose.composer import Composer
from docx import Document

data = {
    "data": [{
        "fio": "test",
        "group": "test_group",
        "practice_name": "test_name_practice"
    }]
}

tpl = DocxTemplate("templates/report.docx")

tpl.render(data)

tpl.save("result/result.docx")
