import os
import win32com.client

class BPac(object):
    def __init__(self):
        # p-PACの場合、`DispatchWithEvents`ではイベント登録できず
        #=> TypeError: This COM object can not automate the makepy process - please run makepy manually for this object
        # from win32com.client import gencache
        # gencache.EnsureModule('{90359D74-B7D9-467F-B938-3883F4CAB582}', 0, 1, 1)
        # self.doc = win32com.client.DispatchWithEvents("bpac.Document", PrintEvents)
        
        # EnsureDispatch()は、エラーになるので使えない
        #=> TypeError: This COM object can not automate the makepy process - please run makepy manually for this object
        # self.doc = win32com.client.gencache.EnsureDispatch("bpac.Document")
        self.doc = win32com.client.DispatchEx("bpac.Document")
        
        # `GetInstalledPrinters()`のように`()`を付けるとエラー
        #=> TypeError: 'tuple' object is not callable
        # printers = self.doc.Printer.GetInstalledPrinters();
        self.printers = self.doc.Printer.GetInstalledPrinters


    def show_printer(self):
        '''
        プリンタの表示
        '''
        for p in self.printers:
            support = "Yes" if self.doc.Printer.IsPrinterSupported(p) else "No"
            status = "Online" if self.doc.Printer.IsPrinterOnline(p) else "Offline"
            print("{name} - Support: {support}, Status: {status}"
                .format(name=p, support=support, status=status))


    def show_media(self):
        '''
        メディア(ラベル)の表示
        '''
        for p in self.printers:
            # 複数プリンタがあってもdoc.Printerだけだと片方しか取得できないため、
            # doc.SetPrinter()で明示的にプリンタを指定してからdoc.Printerで情報を取得する
            # なお、ラベルの調整はしない
            self.doc.SetPrinter(p, False)

            id = self.doc.Printer.GetMediaId
            name = self.doc.Printer.GetMediaName
            msg = "Label - {id} : {name}".format(id=id, name=name) if name else "No Media"
            print(msg)
        
        
    def get_enabled_printer(self):
        '''
        使用できるプリンタの取得
        複数ある場合は、最初のプリンタを取得
        '''
        for p in self.printers:
            self.doc.SetPrinter(p, False)
            
            if  self.doc.Printer.IsPrinterSupported(p) \
            and self.doc.Printer.IsPrinterOnline(p) \
            and self.doc.Printer.GetMediaName:
                return p
        return ""
        
        
    def print_label(self):
        '''
        ラベルの印刷
        '''
        printer = self.get_enabled_printer()
        if not printer:
            print("利用できるプリンタがありません")
            return
        
        input("何かキーを押すと、{}から印刷します。>>>".format(printer))
        
        # 念のため、有効なプリンタを再設定
        self.doc.SetPrinter(printer, False)
        
        # 印刷で使うラベルテンプレートは
        # Pythonスクリプトと同じディレクトリに置く前提
        dir = os.path.abspath(os.path.dirname(__file__))
        lbx_path = os.path.join(dir, "test.lbx")
        
        hasOpened = self.doc.Open(lbx_path)
        if not hasOpened:
            print("指定されたラベルを開けませんでした")
            return
        
        
        # イベントがあるかどうかを調べる
        # b-PAC 3.1だと、以下のような値が返ってくる
        # => <class 'win32com.gen_py.90359D74-B7D9-467F-B938-3883F4CAB582x0x1x3.IPrintEvents'>
        print(win32com.client.getevents("bpac.Document"))

        # pywin32を使ったイベントハンドラの追加
        # VBScript/JScriptのようにSetPrintedCallback()を使うとエラー
        # => TypeError: Objects of type 'type' can not be converted to a COM VARIANT
        # self.doc.SetPrintedCallback(PrintEvents)
        
        # なので、以下の方法で追加する
        # http://d.hatena.ne.jp/yach/20070913
        handler = PrintEvents(self.doc)
        
        
        # カットの設定
        # デフォルトの各ラベルでカットから、最終ラベルでのカットへと変更
        # b-PAC SDKの列挙型を指定した場合はエラーとなるため16進数設定
        # self.doc.StartPrint("", win32com.client.constants.PrintOptionConstants.bpoCutAtEnd)
        #=> AttributeError: PrintOptionConstants
        self.doc.StartPrint("", 0x04000000)

               
        for i in range(0, 3):
            # テンプレートのテキストオブジェクトへの値設定
            self.doc.GetObject("Content").Text = "No.{}".format(i)
            # テンプレートのバーコードオブジェクトへの値設定
            self.doc.SetBarcodeData(self.doc.GetBarcodeIndex("Barcode"), i);
            # 設定終了
            self.doc.PrintOut(1, 0x04000000);
        
        self.doc.EndPrint
        self.doc.Close
        
        # Printedイベント結果を確認するため、少々待つ
        import time
        time.sleep(5)
        
        print("End")
        

# http://d.hatena.ne.jp/yach/20070913
class PrintEvents(win32com.client.getevents('bpac.Document')):
    '''
    Printedイベントのイベントハンドラ用クラス
    イベントハンドラは`On`で始める必要あり：Onがないと動作しない
    '''
    def OnPrinted(self, status, value):
    # def Printed(self, status, value): => 動作しない
        print("status:{s} / value{v}".format(s=status, v=value))


if __name__ == "__main__":
    bpac = BPac()
    
    bpac.show_printer()
    
    bpac.show_media()
    
    bpac.print_label()
    
    
# このスクリプトの実行結果例)
# (env) D:\Sandbox\bpac_python>python pywin32_ver.py
# Brother QL-720NW - Support: Yes, Status: Online
# Label - 259 : 62mm
# 何かキーを押すと、Brother QL-720NWから印刷します。>>>
# <class 'win32com.gen_py.90359D74-B7D9-467F-B938-3883F4CAB582x0x1x3.IPrintEvents'>
# status:0 / value128
# End