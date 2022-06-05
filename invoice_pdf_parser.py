import os
import re
import sys
import traceback

from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LAParams, LTTextBoxHorizontal
from pdfminer.pdfdocument import PDFTextExtractionNotAllowed, PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfparser import PDFParser


def parse_invoice_pdf(invoice_pdf: str) -> dict:
    print("解析文件" + invoice_pdf)
    invoice = {"file": os.path.basename(invoice_pdf)}
    with open(invoice_pdf, "rb") as fp:
        parser = PDFParser(fp)
        doc = PDFDocument(parser)
        parser.set_document(doc)

        if not doc.is_extractable:
            raise PDFTextExtractionNotAllowed

        rsrcmgr = PDFResourceManager()
        # 创建一个PDF设备对象
        laparams = LAParams(boxes_flow=1.5)
        device = PDFPageAggregator(rsrcmgr, laparams=laparams)
        # 创建一个PDF解释其对象
        interpreter = PDFPageInterpreter(rsrcmgr, device)
        money = False
        date = False
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
                if (isinstance(x, LTTextBoxHorizontal)):
                    for obj in x:
                        boxes.append(obj)
            boxes.sort(key=lambda box: (1, box.y0, -box.x0), reverse=True)
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

            for x in boxes:
                text = x.get_text()
                print(text)
                if re.search("小\s*写", text):
                    money = True
                    continue
                elif text.rfind("开票日期") >= 0:
                    date = True
                    continue
                if date and not invoice.get("日期"):
                    year, month, day = text.replace("年", " ").replace("月", " ").replace("日", " ").split()[0:3]
                    invoice["日期"] = "%s年%s月%s日" % (year, month, day)
                    continue
                if money and not invoice.get("金额"):
                    if x.get_text().startswith("¥") or x.get_text().startswith("￥"):
                        invoice["金额"] = float(x.get_text()[len("¥"):])
                    else:
                        invoice["金额"] = float(x.get_text())
                if not invoice.get("类型"):
                    if text.rfind("汽油") > 0 or text.rfind("停车") > 0:
                        invoice["类型"] = "加油费"
                    elif text.rfind("餐饮") > 0:
                        invoice["类型"] = "招待费"
                    elif text.rfind("预付") > 0:
                        invoice["类型"] = "礼品卡"
                    elif text.rfind("客运") > 0 or text.rfind("运输") > 0:
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
