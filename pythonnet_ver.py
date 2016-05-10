import clr
from System import Activator, Type
from System.Reflection import BindingFlags


def main():
    doc = Activator.CreateInstance(Type.GetTypeFromProgID('bpac.Document'))
    print("doc: {}".format(doc.__class__))
    #=> <class 'System.__ComObject'>
    
    # printers = doc.Printer.GetInstalledPrinters
    # => AttributeError: '__ComObject' object has no attribute 'Printer'
    printer = doc.GetType().InvokeMember("Printer", BindingFlags.GetProperty, None, doc, None )
    print("Printer:{}".format(printer))
    
    # 属性を確認
    print(dir(printer))
    
    printers = doc.GetType().InvokeMember("GetInstalledPrinters", BindingFlags.GetProperty, None, printer, None )
    print("InstalledPrinters:{}".format(printers))
    
    for p in printers:
        print("installed Printer:{}".format(p))


def excel():
    xlsx = Activator.CreateInstance(Type.GetTypeFromProgID('Excel.Application'))
    print("xlsx: {}".format(xlsx.__class__))


if __name__ == "__main__":
    main()
    
    excel()