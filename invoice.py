# -*- coding: utf-8 -*-
import os
import threading
import traceback

import wx
import wx.xrc
import wx.grid

from invoice_pdf_parser import parse_invoice_pdf


class MyFileDropTarget(wx.FileDropTarget):
    def __init__(self, window):
        wx.FileDropTarget.__init__(self)
        self._window = window

    def OnDropFiles(self, x, y, filenames):
        self._window.setDirPath(filenames)
        return True


class ParseThread(threading.Thread):
    def __init__(self, window):
        threading.Thread.__init__(self)
        self._window = window

    def run(self):
        self._window.parseInvoices()


class InvoiceParserDlg(wx.Dialog):
    _COLUMNS = ["文件名", "日期", "类型", "金额"]

    def __init__(self, parent):
        wx.Dialog.__init__(self, parent, id=wx.ID_ANY, title="发票解析器(拖拽发票目录到表格即可)", pos=wx.DefaultPosition,
                           size=wx.Size(703, 512), style=wx.DEFAULT_DIALOG_STYLE)

        self.SetSizeHints(wx.DefaultSize, wx.DefaultSize)

        bSizer1 = wx.BoxSizer(wx.VERTICAL)

        bSizer2 = wx.BoxSizer(wx.HORIZONTAL)

        self.m_staticText1 = wx.StaticText(self, wx.ID_ANY, u"发票目录", wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_staticText1.Wrap(-1)

        bSizer2.Add(self.m_staticText1, 0, wx.ALL, 5)

        self.m_dir_ctrl = wx.TextCtrl(self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer2.Add(self.m_dir_ctrl, 0, wx.ALL, 5)

        self.m_browser_button = wx.Button(self, wx.ID_ANY, u"浏览", wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer2.Add(self.m_browser_button, 0, wx.ALL, 5)
        bSizer1.Add(bSizer2, 0, wx.ALL, 5)

        self.m_staticline1 = wx.StaticLine(self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.LI_HORIZONTAL)
        bSizer1.Add(self.m_staticline1, 0, wx.EXPAND | wx.ALL, 5)

        self.m_data = wx.grid.Grid(self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, 0)
        _drag_target = MyFileDropTarget(self)
        self.m_data.SetDropTarget(_drag_target)

        # Grid
        self.m_data.CreateGrid(10, len(self._COLUMNS))
        self.m_data.EnableEditing(False)
        self.m_data.EnableGridLines(True)
        self.m_data.EnableDragGridSize(False)
        self.m_data.SetMargins(0, 0)

        # Columns
        self.m_data.EnableDragColMove(False)
        self.m_data.EnableDragColSize(True)
        self.m_data.SetColLabelSize(30)
        self.m_data.SetColLabelAlignment(wx.ALIGN_CENTER, wx.ALIGN_CENTER)
        for i, v in list(enumerate(self._COLUMNS)):
            self.m_data.SetColLabelValue(i, v)

        # Rows
        self.m_data.EnableDragRowSize(True)
        # self.m_data.SetRowLabelSize(80)
        self.m_data.SetRowLabelAlignment(wx.ALIGN_CENTER, wx.ALIGN_CENTER)

        # Label Appearance

        # Cell Defaults
        self.m_data.SetDefaultCellAlignment(wx.ALIGN_LEFT, wx.ALIGN_TOP)
        bSizer1.Add(self.m_data, 1, wx.ALL | wx.EXPAND, 5)

        self.m_progress = wx.Gauge(self, wx.ID_ANY, 100, wx.DefaultPosition, wx.DefaultSize, wx.GA_HORIZONTAL)
        self.m_progress.SetValue(0)
        bSizer1.Add(self.m_progress, 0, wx.ALL | wx.EXPAND, 5)

        self.SetSizer(bSizer1)
        self.Layout()

        self.Centre(wx.BOTH)
        self.m_browser_button.Bind(wx.EVT_BUTTON, self.onBrowse)
        self.Bind(wx.EVT_CLOSE, self.onClose)

        self._paths = []
        self._parseThread = None

    def __del__(self):
        pass

    def onClose(self, event):
        if self._parseThread:
            self._parseThread.join()
        self.Destroy()

    def onBrowse(self, event):
        with wx.DirDialog(self, message="请选择目录或文件", style=wx.FD_OPEN) as dirDialog:
            ret = dirDialog.ShowModal()
            if ret == wx.ID_CANCEL or not dirDialog.GetPath():
                return
            self.setDirPath(dirDialog.GetPath())

    def setDirPath(self, dir):
        if dir:
            if isinstance(dir, list):
                self._paths = dir
            else:
                self._paths = [dir]
            self.m_dir_ctrl.SetValue(','.join(self._paths))
        self.updateGrid()

    def parseInvoices(self):
        if not self._paths:
            return
        files = []
        for path in self._paths:
            if os.path.isdir(path):
                for f in os.listdir(path):
                    if f.lower().endswith("pdf"):
                        files.append(os.path.join(path, f))
            elif path.lower().endswith("pdf"):
                files.append(path)
        money_sum = 0
        row = 0
        for invoice_pdf in files:
            if row == self.m_data.GetNumberRows():
                self.m_data.AppendRows(1)
            try:
                invoice = parse_invoice_pdf(invoice_pdf)
            except:
                traceback.print_exc()
                continue
            for i, v in list(enumerate(self._COLUMNS)):
                value = invoice.get(v, ' ')
                if v == "金额":
                    value = "¥%s" % value
                self.m_data.SetCellValue(row, i, str(value))
            money_sum += invoice.get('金额', 0.)
            row += 1
            self.m_progress.SetValue(row * 100 // len(files))

        if row == self.m_data.GetNumberRows():
            self.m_data.AppendRows(1)
        self.m_data.SetCellValue(row, 0, "合计")
        self.m_data.SetCellValue(row, 3, str(money_sum))
        return

    def updateGrid(self):
        self.m_progress.SetValue(0)
        self.m_data.ClearGrid()
        self._parseThread = ParseThread(self)
        self._parseThread.start()


def main():
    root = wx.App(0)
    dialog = InvoiceParserDlg(parent=None)
    dialog.Show()
    root.MainLoop()


if __name__ == "__main__":
    main()
