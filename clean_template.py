"""打包用：把 template.docx 的個人內容換成範例文字，存到 _dist_template\template.docx。
保留表格格式與編號段落（report.build 靠 (2,1)、(2,2)、(4,1) 的段落當模子）。"""
from pathlib import Path

import docx

import report as R

SRC = R.BASE / "template.docx"
OUT = R.BASE / "_dist_template" / "template.docx"


def main():
    d = docx.Document(SRC)
    rows = d.tables[0].rows
    R._set_text(rows[0].cells[0].paragraphs[0]._p, "姓名 職稱")
    for r in list(rows)[1:]:
        for i, c in enumerate(r.cells):
            for p in c.paragraphs:
                if p.runs:
                    R._set_text(p._p, "專案名稱" if i == 0 else "（範例）")
    d.core_properties.author = d.core_properties.last_modified_by = ""
    OUT.parent.mkdir(exist_ok=True)
    d.save(OUT)
    print(OUT)


if __name__ == "__main__":
    main()
