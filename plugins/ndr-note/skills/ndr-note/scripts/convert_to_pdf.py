# -*- coding: utf-8 -*-
"""
docx -> pdf 변환 (OS 독립).
우선순위: 1) LibreOffice(soffice/libreoffice, 전 OS)  2) Windows + MS Word(COM)
사용법: python convert_to_pdf.py "<파일.docx>"
"""
import os, sys, shutil, subprocess

def to_pdf(docx_path):
    docx_path = os.path.abspath(docx_path)
    if not os.path.isfile(docx_path):
        raise FileNotFoundError(docx_path)
    out_dir = os.path.dirname(docx_path)
    pdf_path = os.path.splitext(docx_path)[0] + ".pdf"

    # 1) LibreOffice (Windows/Mac/Linux 공통)
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice:
        subprocess.run([soffice, "--headless", "--convert-to", "pdf",
                        "--outdir", out_dir, docx_path], check=True)
        return pdf_path

    # 2) Windows + MS Word (COM via PowerShell)
    if os.name == "nt":
        ps = (
            "$w=New-Object -ComObject Word.Application;$w.Visible=$false;"
            "$d=$w.Documents.Open('%s');"
            "$d.SaveAs([ref]'%s',[ref]17);$d.Close();$w.Quit()"
            % (docx_path.replace("'", "''"), pdf_path.replace("'", "''"))
        )
        subprocess.run(["powershell", "-NoProfile", "-Command", ps], check=True)
        return pdf_path

    raise RuntimeError(
        "PDF 변환 도구가 없습니다. LibreOffice(soffice)를 설치하거나 Windows+MS Word 환경에서 실행하세요."
    )

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print('사용법: python convert_to_pdf.py "<파일.docx>"'); sys.exit(1)
    print("PDF:", to_pdf(sys.argv[1]))
