# /// script
# requires-python = ">=3.10"
# dependencies = ["python-docx", "lxml"]
# ///
"""Imprime el XML (resumido) de una tabla de un docx. Uso: dump_table_xml.py <docx> <indice_tabla>"""
import re
import sys

from docx import Document
from lxml import etree

sys.stdout.reconfigure(encoding="utf8")
d = Document(sys.argv[1])
xml = etree.tostring(d.tables[int(sys.argv[2])]._tbl, pretty_print=True, encoding="unicode")
xml = re.sub(r' xmlns:\w+="[^"]+"', "", xml)
print(xml[:9000])
