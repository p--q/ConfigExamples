#!/opt/libreoffice5.2/program/python
# -*- coding: utf-8 -*-
import officehelper
import traceback
from functools import wraps
import sys
from com.sun.star.beans import PropertyValue
def main(ctx, smgr):
    cp = smgr.createInstanceWithContext("com.sun.star.configuration.ConfigurationProvider", ctx)
    print("\nStarting examples.")
    readDataExample(cp)



def readDataExample(cp):
    print("\n--- starting example: read grid option settings --------------------")
    data = readGridConfiguration(cp)
#     print("Read grid options: " + data)
def readGridConfiguration(cp):
    node = PropertyValue(Name="nodepath", Value="/org.openoffice.Office.Calc/Grid")
    ca = cp.createInstanceWithArguments("com.sun.star.configuration.ConfigurationAccess", (node,))
    sub = ca.getHierarchicalPropertyValue("Subdivision")
    elems = sub.getPropertyValues(("XAxis", "YAxis"))
    GridOptions.subdivision_x = int(elems[0])
    GridOptions.subdivision_y = int(elems[1])
    
    print(vars(GridOptions))
    
    return GridOptions
#     def createConfigurationView(ph):
#         prop = PropertyValue(Name="nodepath", Value=ph)
#         return cp.createInstanceWithArguments("com.sun.star.configuration.ConfigurationAccess", (prop,))



class GridOptions():
    pass
        



# funcの前後でOffice接続の処理
def connectOffice(func):
    @wraps(func)
    def wrapper():  # LibreOfficeをバックグラウンドで起動してコンポーネントテクストとサービスマネジャーを取得する。
        ctx = None
        try:
            ctx = officehelper.bootstrap()  # コンポーネントコンテクストの取得。
        except:
            pass
        if not ctx:
            print("Could not establish a connection with a running office.")
            sys.exit()
        print("Connected to a running office ...")
        smgr = ctx.getServiceManager()  # サービスマネジャーの取得。
        if not smgr:
            print( "ERROR: no service manager" )
            sys.exit()
        print("Using remote servicemanager\n") 
        try:
            func(ctx, smgr)  # 引数の関数の実行。
        except:
            traceback.print_exc()
        # soffice.binの終了処理。これをしないとLibreOfficeを起動できなくなる。
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        prop = PropertyValue(Name="Hidden",Value=True)
        desktop.loadComponentFromURL("private:factory/swriter", "_blank", 0, (prop,))  # バックグラウンドでWriterのドキュメントを開く。
        terminated = desktop.terminate()  # LibreOfficeをデスクトップに展開していない時はエラーになる。
        if terminated:
            print("\nThe Office has been terminated.")  # 未保存のドキュメントがないとき。
        else:
            print("\nThe Office is still running. Someone else prevents termination.")  # 未保存のドキュメントがあってキャンセルボタンが押された時。
    return wrapper
if __name__ == "__main__":
    main = connectOffice(main)
    main()