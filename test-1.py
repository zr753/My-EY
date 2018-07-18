# coding=utf-8
import sys
import importlib
import json
import uniout
import re


from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LTTextBoxHorizontal,LAParams
from pdfminer.pdfdocument import PDFTextExtractionNotAllowed


def parse(path):
    fp = open(path, 'rb') # 以二进制读模式打开
    #用文件对象来创建一个pdf文档分析器
    praser = PDFParser(fp)
    # 创建一个PDF文档
    doc = PDFDocument(praser)
    # 连接分析器 与文档对象
    praser.set_document(doc)
    #doc.set_parser(praser)

    outlines = doc.get_outlines()
    for (level,title,dest,a,se) in outlines:
        print (level, json.dumps(title, ensure_ascii=False))

if __name__ == '__main__':
    path =r'C:/Users/raymond-r.zhang/Desktop/Python Scripts/capital_commitment/000895.pdf'
    parse(path)
