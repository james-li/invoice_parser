import os
import re
import sys
import traceback
from collections import OrderedDict

from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LAParams, LTTextBoxHorizontal
from pdfminer.pdfdocument import PDFTextExtractionNotAllowed, PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfparser import PDFParser


def cidToChar(cidx):
    return chr(int(re.findall(r'\(cid\:(\d+)\)', cidx)[0]) + 29)


def append_lines(lines: OrderedDict, boxes: list, page_no: int = 0):
    if lines:
        last_key = next(reversed(lines))
    else:
        last_key = 0
    for box in boxes:
        x0, y0 = box.bbox[0], box.bbox[1]
        if abs(y0 - last_key) > 2:
            last_key = y0
        line = lines.get(last_key, [])
        line.append(box)
        line.sort(key=lambda x: x.x0)
        lines[last_key] = line


def get_pdf_content_lines(doc: PDFDocument) -> OrderedDict:
    lines = OrderedDict()
    if not doc.is_extractable:
        raise PDFTextExtractionNotAllowed

    rsrcmgr = PDFResourceManager()
    # 创建一个PDF设备对象
    laparams = LAParams(boxes_flow=1.5)
    device = PDFPageAggregator(rsrcmgr, laparams=laparams)
    # 创建一个PDF解释其对象
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    # 循环遍历列表，每次处理一个page内容
    # doc.get_pages() 获取page列表
    for page_no, page in enumerate(PDFPage.create_pages(doc)):
        interpreter.process_page(page)
        # 接受该页面的LTPage对象
        layout = device.get_result()
        # 这里layout是一个LTPage对象 里面存放着 这个page解析出的各种对象
        # 一般包括LTTextBox, LTFigure, LTImage, LTTextBoxHorizontal 等等
        # 想要获取文本就获得对象的text属性，
        boxes = []
        for x in layout:
            if isinstance(x, LTTextBoxHorizontal):
                for obj in x:
                    boxes.append(obj)
        boxes.sort(key=lambda box: (1, box.bbox[1], -box.bbox[0]), reverse=True)
        obj0 = None
        for x in boxes:
            if not obj0:
                obj0 = x
                continue
            obj1 = x
            if obj1.is_voverlap(obj0) and min(obj0.height, obj1.height) * laparams.line_overlap < obj0.voverlap(
                    obj1):
                obj1.y0 = obj0.y0
            obj0 = obj1
        boxes.sort(key=lambda box: (1, box.y0, -box.x0), reverse=True)
        append_lines(lines, boxes, page_no)

    return lines


def parse_invoice_pdf(invoice_pdf: str) -> dict:
    print("解析文件" + invoice_pdf)
    invoice = {"文件名": os.path.basename(invoice_pdf)}
    with open(invoice_pdf, "rb") as fp:
        parser = PDFParser(fp)
        doc = PDFDocument(parser)
        parser.set_document(doc)
        lines = get_pdf_content_lines(doc)
        money = False
        date = False
        for x in lines.values():
            text = " ".join([box.get_text() for box in x])
            if text.rfind('(cid') >= 0:
                tl = text.replace(' ', '').replace(')(cid:', ',').replace('(cid:', '').replace(')', '').split(',')
                text = ''.join([chr(int(i) + 29) for i in tl])
            print(text)
            if text.rfind("开票日期") >= 0:
                if not date:
                    try:
                        year, month, day = re.findall(r"\d+", text)[0:3]
                        invoice["日期"] = "%s年%s月%s日" % (year, month, day)
                        date = True
                    except:
                        continue
            elif re.search("小\s*写", text):
                if not money:
                    invoice["金额"] = float(text.split(' ')[-1].replace('￥', '').replace('¥', ''))
                    money = True
            if not invoice.get("类型"):
                if text.rfind("汽油") > 0:
                    invoice["类型"] = "加油费"
                elif text.rfind("餐饮") > 0 or text.rfind("咖啡") > 0:
                    invoice["类型"] = "招待费"
                elif text.rfind("预付") > 0:
                    invoice["类型"] = "礼品卡"
                elif text.rfind("客运") > 0 or text.rfind("运输") > 0 or text.rfind("停车") > 0:
                    invoice["类型"] = "交通费"

    return invoice


if __name__ == "__main__":
    dir = sys.argv[1]
    money = []
    for entry in os.listdir(dir):
        if entry.endswith(".pdf"):
            try:
                money.append(parse_invoice_pdf(os.path.join(dir, entry)).get("money", 0))
            except:
                traceback.print_exc()

    print(money)
    print(sum(money))
