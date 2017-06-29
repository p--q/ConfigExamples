#!/opt/libreoffice5.2/program/python
# -*- coding: utf-8 -*-
import unohelper
import officehelper
import traceback
from functools import wraps
import sys
from com.sun.star.beans import PropertyValue
from collections import namedtuple
from com.sun.star.uno import RuntimeException
from com.sun.star.util import XChangesListener
import types


def main(ctx, smgr):  # ctx: コンポーネントコンテクスト、smgr: サービスマネジャー
    cp = createProvider(ctx, smgr)  # ConfigurationProviderの取得。
    if checkProvider(cp):
        print("\nStarting examples.")
        readDataExample(cp)  # /org.openoffice.Office.Calc/Grid以下の特定の値を取得する例。
        browseDataExample(cp)  # /org.openoffice.TypeDetection.Filter/Filters以下の値一覧を出力する例。
        updateGroupExample(cp)  # /org.openoffice.Office.Calc/Grid以下の値を変更する例。
#         resetGroupExample(cp)  # デフォルト値に戻す例。動きません。
        print("\nAll Examples completed.")
    else:
        print("ERROR: Cannot run examples without ConfigurationProvider.")
def createProvider(ctx, smgr):  # ConfigurationProviderをインスタンス化。引数なしでインスタンス化しているのでDefaultProviderが返る。
    return smgr.createInstanceWithContext("com.sun.star.configuration.ConfigurationProvider", ctx)
def checkProvider(cp):  # ConfigurationProviderの情報を取得。
    if cp is None:
        print("No provider available. Cannot access configuration data.")
        return False
    try:
        if not cp.supportsService("com.sun.star.configuration.ConfigurationProvider"):  # com.sun.star.configuration.ConfigurationProviderサービスをサポートしていないとき    
            print("WARNING: The provider is not a 'com.sun.star.configuration.ConfigurationProvider'") 
        services = cp.getSupportedServiceNames()  # 取得したConfigurationProviderがサポートするサービスを取得。
        t = ("a ", str(services).strip("(),"), "") if len(services)==1 else ("", str(services).strip("()"), "s")  # 複数形への対応。
        print("The provider has {}{} service{}.".format(*t))  # 取得したConfigurationProviderがサポートするサービス一覧を出力。
        print("Using provider implementation: {}.".format(cp.getImplementationName()))  #  ConfigurationProviderの実装名を出力。
        return True
    except RuntimeException:
        print("ERROR: Failure while checking the provider services.")
        traceback.print_exc()
        return False
    
    
def readDataExample(cp):  # /org.openoffice.Office.Calc/Grid以下の特定の値を取得する例。
    try:
        print("\n--- starting example: read grid option settings --------------------")
        options = readGridConfiguration(cp)  # namedtupleを受け取る。
        print("Read grid options: {}".format(options))
    except:
        traceback.print_exc()
def readGridConfiguration(cp):  # 設定ファイルの読み込み
    configreader = createConfigReader(cp)  # 読み込み専用の関数を取得。
    root = configreader("/org.openoffice.Office.Calc/Grid")  # 引数のパスで根ノードを取得。
    root = Proxy(root)  # ノードにgetNodeメソッドを付加。
    visible = root.getNode("Option/VisibleGrid")  # サブノードOption/VisibleGridの値を取得。
    resolution_x, resolution_y = root.getNode("Resolution").getNode("XAxis/Metric", "YAxis/Metric")  # サブノードResolutionのサブノードXAxis/MetricとYAxis/Metricの値を取得。
    subdivision_x, subdivision_y = root.getNode("Subdivision").getNode("XAxis", "YAxis")  # サブノードSubdivisionのサブノードXAxisとYAxisの値を取得。
    root.dispose()  # ConfigurationAccessサービスのインスタンスを破棄。
    return GridOptions(visible, resolution_x, resolution_y, subdivision_x, subdivision_y)  # namedtupleを返す。
def createConfigReader(cp):  # ConfigurationProviderサービスのインスタンスを受け取る高階関数。
    def getRoot(path):  # ConfigurationAccessサービスのインスタンスを返す関数。
        node = PropertyValue(Name="nodepath", Value=path)
        return cp.createInstanceWithArguments("com.sun.star.configuration.ConfigurationAccess", (node,))
    return getRoot
class Proxy:  # Proxyパターンでインスタンスにメソッドを追加する。
    def __init__(self, obj):  # メソッドを追加するインスタンスを取得。
        self._obj = obj
    def getNode(self, *args):  # インスタンスに追加するメソッド。
        delimset = {"/", ".", ":"}  # パス区切り一覧
        if len(args)==1:  # 引数の数が1つのとき
            node = self._obj.getHierarchicalPropertyValue(*args) if delimset & set(*args) else self._obj.getPropertyValue(*args)  # パス区切りの有無でgetHierarchicalPropertyValue()とgetPropertyValue()を使い分ける。
            return Proxy(node) if type(node).__name__=="pyuno" else node  # nodeがPyUNOオブジェクトのときはProxyクラスのインスタンスを返し、そうでないときはそのまま返す。
        elif len(args)>1:  # 引数の数が2つ以上のとき
            nodes = self._obj.getHierarchicalPropertyValues(args) if delimset & set("".join(args)) else self._obj.getPropertyValues(args)  # パス区切りの有無でgetHierarchicalPropertyValues()とgetPropertyValues()を使い分ける。
            return [Proxy(node) if type(node).__name__=="pyuno" else node for node in nodes]  # 各ノードについてPyUNOオブジェクトのときはProxyクラスのインスタンスを、そうでないときはそのままを要素にしたリストを返す。
    def __getattr__(self, name):  # Proxyクラス属性にnameが見つからなかったときにnameを引数にして呼び出されます。__setattr__()や __delattr__()が常に呼び出されるのとは対照的です。
        return getattr(self._obj, name)  # Proxyクラスのインスタンスが取得したインスタンスの属性としてnameを呼び出す。
    def __setattr__(self, name, value):  # アンダースコアが始まる属性名のときはProxyの属性にvalueを代入し、そうでない時はProxyクラスのインスタンスが取得したインスタンスの属性にvalueを代入する。
        super().__setattr__(name, value) if name.startswith('_') else setattr(self._obj, name, value)
    def __delattr__(self, name):  # アンダースコアが始まる属性名のときはProxyの属性を削除し、そうでない時はProxyクラスのインスタンスが取得したインスタンスの属性を削除する。
        super().__delattr__(name) if name.startswith('_') else delattr(self._obj, name)   
class GridOptions(namedtuple("GridOptions", "visible resolution_x resolution_y subdivision_x subdivision_y")):  # namedtupleの__str__()メソッドを上書きする。
    __slots__ = ()  # インスタンス辞書の作成抑制。
    def __str__(self):  # 文字列として呼ばれた場合に返す値を設定。
        return "[ Grid is {0}; resolution = ({1},{2}); subdivision = ({3},{4}) ]"\
            .format("VISIBLE" if self.visible else "HIDDEN", self.resolution_x, self.resolution_y, self.subdivision_x, self.subdivision_y)


def browseDataExample(cp):  # /org.openoffice.TypeDetection.Filter/Filters以下の値一覧を出力する例。
    try:
        print("\n--- starting example: browse filter configuration ------------------")
        printRegisteredFilters(cp)
    except:
        traceback.print_exc()
def printRegisteredFilters(cp):
    configreader = createConfigReader(cp)  # 読み込み専用の関数を取得。
    root = configreader("/org.openoffice.TypeDetection.Filter/Filters")  # 引数のパスで根ノードを取得。
    e = Evaluator()  # Visitorパターンをインスタンス化。
    output = e.visit(root)  # VisitorパターンでCompositeパターンに出力機能を追加。リストを取得。
    print("\n".join(output))  # リストの要素を改行して出力。
    root.dispose()  # ConfigurationAccessサービスのインスタンスを破棄。
class Visit:  # ノードを選別するためのクラス。
    def __init__(self, node):
        self.node = node   
class NodeVisitor:  # ジェネレーター版Vistorパターン
    def visit(self, node):
        stack = [Visit(node)]  # ノードをVisitクラスのインスタンスにする。
        last_result = []  # 結果を入れるリスト。
        while stack:  # スタックがある間実行。
            try:
                last = stack[-1]  # スタックの最後の要素を取得。
                if isinstance(last, types.GeneratorType):  # lastがジェネレーターのとき
                    stack.append(next(last))  # ジェネレーターから次の値を取得。
                elif isinstance(last, Visit):  # lastがVisitのインスタンスのとき
                    stack.append(self._visit(stack.pop().node))  # スタックの最後の値を取り出してノードを_visitメソッドに渡した戻り値をスタックに取得。
                else:
                    last_result.append(stack.pop())  # lastがジェネレーターでもVisitのインスタンスでもないときはstackから取り出してlast_resultの要素に追加する。
            except StopIteration:  # ジェネレーターから値が取得できなかったとき
                stack.pop()  # ジェネレーターを捨てる。
        return last_result  # 結果を取得したリストを返す。
    def _visit(self, node):  # 各ノードでの処理を振り分ける。
        name = "PyUNO" if type(node).__name__=="pyuno" else "Values"  # ノードがPyUNOオブジェクトかそうでないかで振り分け。
        self.methname = 'visit_{}'.format(name)  # ノードに適用するメソッド名を取得。
        meth = getattr(self, self.methname, None)  # selfの属性にあるメソッドを取得。メソッドが存在しないときはNoneを返す。
        if meth is None:  # メッソドが存在しなかったとき
            meth = self.generic_visit  # generic_visit()メソッドを取得。
        return meth(node)  # 引数をノードにしてメソッドを返す。
    def generic_visit(self, node):  # 存在しないメソッドが呼ばれた時に呼ばれるメソッド。
        raise RuntimeError('No {} method'.format(self.methname))
class Evaluator(NodeVisitor):  # ノードに適用するメソッドを持つNodeVisitorのサブクラス。これらのメソッドはジェネレーター。
    def visit_Values(self, node):  # ノードがPyUNOオブジェクト以外の時
        if isinstance(node, tuple):  # タプルの時
            yield "\tValue: {0} = {{{1}}}".format(self.path, ", ".join(node))
        else:  # タプルでない時
            yield "\tValue: {} = {}".format(self.path, node)
    def visit_PyUNO(self, node):  # ノードがPyUNOオブジェクトのとき
        if hasattr(node, "getTemplateName") and node.getTemplateName().endswith("Filter"):
            yield "Filter {} ({})".format(node.getName(), node.getHierarchicalName())
        childnames = node.getElementNames()
        for childname in childnames:  # サブノードについて
            self.path = node.composeHierarchicalName(childname)
            yield Visit(node.getByName(childname))  # Evaluatorのメソッドで処理するためにVisitクラスのインスタンスにして返す。

            
def updateGroupExample(cp):  # /org.openoffice.Office.Calc/Grid以下の値を変更する例。
    try:
        print("\n--- starting example: update group data --------------")
        editGridOptions(cp)
    except:
        traceback.print_exc()
def editGridOptions(cp):
    config = createConfigUpdater(cp)  # 読み書き用の関数を取得。
    path = "/org.openoffice.Office.Calc/Grid"
    model = config(path)  # 引数のパスで根ノードをモデルとして取得。
    controller = GridOptionsEditor(model)  # モデルを引数にしてコントローラを取得。
    controller.changeSomeData(config(path + "/Subdivision"))  # コントローラでモデルを変更する。
    if controller.execute()==GridOptionsEditor.SAVE_SETTINGS:  # さらにモデルを変更する。
        try:
            model.commitChanges()  # モデルの変更を書き込む。
        except Exception as e:
            controller.informUserOfError(e)        
    model.dispose()  # モデルを破棄する。
def createConfigUpdater(cp):
    def getRoot(path):  # ConfigurationUpdateAccessサービスのインスタンスを返す。
        node = PropertyValue(Name="nodepath", Value=path)
        return cp.createInstanceWithArguments("com.sun.star.configuration.ConfigurationUpdateAccess", (node,))
    return getRoot
class GridOptionsEditor:  # コントローラ
    CANCELED = 0
    SAVE_SETTINGS = 1
    def __init__(self, model):
        self.model = model  # モデルを取得。
        self.view = GridOptionsEditorView(model)  # ビューを取得
    def execute(self):  # 孫ノードの値を変更する例の成否を返す。
        try:
            print("-- GridEditor executing --")
            self.toggleVisibility()
            print("-- GridEditor done      --")
            return self.SAVE_SETTINGS
        except Exception as e:
            self.informUserOfError(e)
            return self.CANCELED
    @staticmethod
    def informUserOfError(e):  # 例外のときに実行するコード。インスタンスに関係なく呼び出すのでスタティックメソッドにしている。
        print("ERROR in GridEditor:")
        traceback.print_exc()
    def toggleVisibility(self): # 孫ノードの値を変更する例
        try:
            setting = "Option/VisibleGrid"
            print("GridEditor: toggling Visibility")
            oldval = self.model.getHierarchicalPropertyValue(setting)  # getByHierarchicalName()メソッドに置換可能。
            newval = False if oldval else True
            self.model.setHierarchicalPropertyValue(setting, newval)  # この実行後にXChangesListenerが呼び出される。replaceByHierarchicalName()メソッドに置換可能。
        except Exception as e:
            self.informUserOfError(e)    
    def changeSomeData(self, root):  # 子ノードの値を変更する例。
        try:          
            itemnames = root.getElementNames()
            for itemname in itemnames:
                item = root.getByName(itemname)  # getPropertyValue()メソッドで置換可能。
                if isinstance(item, bool):
                    print("Replacing boolean value: {}".format(itemname))
                    root.replaceByName(itemname, False if item else True)  # setPropertyValue()メソッドで置換可能。
                elif isinstance(item, int):
                    item = 9999-item
                    print("Replacing integer value: {}".format(itemname))
                    root.replaceByName(itemname, item)  # setPropertyValue()メソッドで置換可能。
            root.commitChanges()  # この実行後にXChangesListenerが呼び出される。
            root.dispose()
        except:
            print("Could not change some data in a different view. An exception occurred:")
            traceback.print_exc()     
class GridOptionsEditorView:  # ビュー
    def __init__(self, model):
        self.model = model  # モデルを取得。
        self.createChangesListener()  # モデルにリスナーを付ける。
        self.updateView()  # ビューを更新。    
    def updateView(self):
        if self.model is not None:
            print("Grid options editor: data={}".format(self.readModel()))
        else:
            print("Grid options editor: no model set")
    def readModel(self):  # モデルの情報をnamedtupleに入れて返す。
        try:
            options = "Option/VisibleGrid", "Resolution/XAxis/Metric", "Resolution/YAxis/Metric", "Subdivision/XAxis", "Subdivision/YAxis"
            values = self.model.getHierarchicalPropertyValues(options)
            return  GridOptions(*values)
        except Exception as e:
            GridOptionsEditor.informUserOfError(e)
            return None 
    def createChangesListener(self):  # リスナーをモデルに付ける。
        self.model.addChangesListener(ChangesListener(self))
class ChangesListener(unohelper.Base, XChangesListener):  # モデルに付けるリスナー。      
    def __init__(self, cast):
        self.cast = cast                   
    def changesOccurred(self, event):  # 子ノードのときはcommitChanges()したとき、孫ノードのときは変更した時点で呼び出される。
        print("GridEditor - Listener received changes event containing {} change(s).".format(len(event.Changes)))
        self.cast.updateView()
    def disposing(self, source):  # パブリシャでdispose()したあとに呼ばれる。呼ばれるときにパブリシャは消滅済。
        print("GridEditor - Listener received disposed event: releasing model")


def resetGroupExample(cp):  # デフォルト値に戻す例。UNOIDL未実装のため動かない。
    try:
        print("\n--- starting example: reset group data -----------------------------")
        olddata = readGridConfiguration(cp)
        resetGridConfiguration(cp)
        newdata = readGridConfiguration(cp)
        print("Before reset:   user grid options: {}".format(olddata))
        print("After reset: default grid options: {}".format(newdata))
    except:
        traceback.print_exc()     
def resetGridConfiguration(cp):
    config = createConfigUpdater(cp)
    path = "/org.openoffice.Office.Calc/Grid/Option"
    model = config(path)
    model.setPropertyToDefault("VisibleGrid")  # setPropertyToDefault()メソッドは実装されておらず動きません。


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