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


def main(ctx, smgr):
    cp = createProvider(ctx, smgr)
    if checkProvider(cp):
        print("\nStarting examples.")
#         readDataExample(cp)
#         browseDataExample(cp)
        updateGroupExample(cp)
#         resetGroupExample(cp)
        print("\nAll Examples completed.")
    else:
        print("ERROR: Cannot run examples without ConfigurationProvider.")
def createProvider(ctx, smgr):
    return smgr.createInstanceWithContext("com.sun.star.configuration.ConfigurationProvider", ctx)
def checkProvider(cp):
    if cp is None:
        print("No provider available. Cannot access configuration data.")
        return False
    try:
        if not cp.supportsService("com.sun.star.configuration.ConfigurationProvider"):      
            print("WARNING: The provider is not a 'com.sun.star.configuration.ConfigurationProvider'") 
        services = cp.getSupportedServiceNames()
        t = ("a ", str(services).strip("(),"), "") if len(services)==1 else ("", str(services).strip("()"), "s")
        print("The provider has {}{} service{}.".format(*t))
        print("Using provider implementation: {}.".format(cp.getImplementationName()))
        return True
    except RuntimeException:
        print("ERROR: Failure while checking the provider services.")
        traceback.print_exc()
        return False
    
    
def readDataExample(cp):
    try:
        print("\n--- starting example: read grid option settings --------------------")
        options = readGridConfiguration(cp)
        print("Read grid options: {}".format(options))
    except:
        traceback.print_exc()
class Proxy:
    def __init__(self, obj):
        self._obj = obj
    def getNode(self, *args):
        delimset = {"/", ".", ":"}
        if len(args)==1:
            node = self._obj.getHierarchicalPropertyValue(*args) if delimset & set(*args) else self._obj.getPropertyValue(*args)
            return Proxy(node) if hasattr(node, "getTypes") else node
        elif len(args)>1:
            nodes = self._obj.getHierarchicalPropertyValues(args) if delimset & set("".join(args)) else self._obj.getPropertyValues(args)
            return [Proxy(node) if hasattr(node, "getTypes") else node for node in nodes]
    def __getattr__(self, name):
        return getattr(self._obj, name)
    def __setattr__(self, name, value):
        super().__setattr__(name, value) if name.startswith('_') else setattr(self._obj, name, value)
    def __delattr__(self, name):
        super().__delattr__(name) if name.startswith('_') else delattr(self._obj, name)   
def readGridConfiguration(cp):
    ca = createConfigurationView("/org.openoffice.Office.Calc/Grid", cp)
    root = Proxy(ca)
    visible = root.getNode("Option/VisibleGrid")
    resolution_x, resolution_y = root.getNode("Resolution").getNode("XAxis/Metric", "YAxis/Metric")
    subdivision_x, subdivision_y = root.getNode("Subdivision").getNode("XAxis", "YAxis")
    ca.dispose()
    return GridOptions(visible, resolution_x, resolution_y, subdivision_x, subdivision_y) 
def createConfigurationView(path, cp):
    node = PropertyValue(Name="nodepath", Value=path)
    return cp.createInstanceWithArguments("com.sun.star.configuration.ConfigurationAccess", (node,))
class GridOptions(namedtuple("GridOptions", "visible resolution_x resolution_y subdivision_x subdivision_y")):
    __slots__ = ()
    def __str__(self):
        return "[ Grid is {0}; resolution = ({1},{2}); subdivision = ({3},{4}) ]"\
            .format("VISIBLE" if self.visible else "HIDDEN", self.resolution_x, self.resolution_y, self.subdivision_x, self.subdivision_y)


def browseDataExample(cp):
    try:
        print("\n--- starting example: browse filter configuration ------------------")
        printRegisteredFilters(cp)
    except:
        traceback.print_exc()
def printRegisteredFilters(cp):
    path = "/org.openoffice.TypeDetection.Filter/Filters"
    ca = createConfigurationView(path, cp)
    e = Evaluator()
    output = e.visit(ca)
    print("\n".join(output))
    ca.dispose()
class Visit:
    def __init__(self, node):
        self.node = node   
class NodeVisitor:
    def visit(self, node):
        stack = [Visit(node)]
        last_result = []
        while stack:
            try:
                last = stack[-1]
                if isinstance(last, types.GeneratorType):
                    stack.append(next(last))
                elif isinstance(last, Visit):
                    stack.append(self._visit(stack.pop().node))
                else:
                    last_result.append(stack.pop())
            except StopIteration:
                stack.pop()
        return last_result
    def _visit(self, node):
        name = "PyUNO" if type(node).__name__=="pyuno" else "Values"
        self.methname = 'visit_{}'.format(name)
        meth = getattr(self, self.methname, None)
        if meth is None:
            meth = self.generic_visit
        return meth(node)
    def generic_visit(self, node):
        raise RuntimeError('No {} method'.format(self.methname))
class Evaluator(NodeVisitor):
    def visit_Values(self, node):
        if isinstance(node, tuple):
            yield "\tValue: {0} = {{{1}}}".format(self.path, ", ".join(node))
        else:
            yield "\tValue: {} = {}".format(self.path, node)
    def visit_PyUNO(self, node):
        if hasattr(node, "getTemplateName") and node.getTemplateName().endswith("Filter"):
            yield "Filter {} ({})".format(node.getName(), node.getHierarchicalName())
        childnames = node.getElementNames()
        for childname in childnames:
            self.path = node.composeHierarchicalName(childname)
            yield Visit(node.getByName(childname))

            
def updateGroupExample(cp):
    try:
        print("\n--- starting example: update group data --------------")
        editGridOptions(cp)
    except:
        traceback.print_exc()
def editGridOptions(cp):
    path = "/org.openoffice.Office.Calc/Grid"
    model = getUpdatableModel(path, cp)
#     view = GridOptionsEditorView(model)
    controller = GridOptionsEditor(model)
    changeSomeData("{}/Subdivision".format(path), cp)
#     if controller.execute()==GridOptionsEditor.SAVE_SETTINGS:
#         try:
#             model.commitChanges()
#         except Exception as e:
#             controller.informUserOfError(e)          
    model.dispose()                   
def getUpdatableModel(path, cp):
    node = PropertyValue(Name="nodepath", Value=path)
    return cp.createInstanceWithArguments("com.sun.star.configuration.ConfigurationUpdateAccess", (node,))
class GridOptionsEditor:
    CANCELED = 0
    SAVE_SETTINGS = 1
    def __init__(self, model):
        self.model = model
        self.view = GridOptionsEditorView(model)
    def execute(self):
        try:
            print("-- GridEditor executing --")
            self.toggleVisibility()
            print("-- GridEditor done      --")
            return self.SAVE_SETTINGS
        except Exception as e:
            self.informUserOfError(e)
            return self.CANCELED
    @staticmethod
    def informUserOfError(e):
        print("ERROR in GridEditor:")
        traceback.print_exc()
    def toggleVisibility(self):
        try:
            setting = "Option/VisibleGrid"
            print("GridEditor: toggling Visibility")
            oldval = self.model.getHierarchicalPropertyValue(setting)
            newval = False if oldval else True
            self.model.setHierarchicalPropertyValue(setting, newval)
        except Exception as e:
            self.informUserOfError(e)      
def changeSomeData(path, cp):
    try:
        model = getUpdatableModel(path, cp)
        itemnames = model.getElementNames()
        for itemname in itemnames:
            item = model.getByName(itemname)
            if isinstance(item, bool):
                print("Replacing boolean value: {}".format(itemname))
                model.replaceByName(itemname, False if item else True)
            elif isinstance(item, int):
                item = 9999-item
                print("Replacing integer value: {}".format(itemname))
                model.replaceByName(itemname, item)
        model.commitChanges()
        model.dispose()
    except:
        print("Could not change some data in a different view. An exception occurred:")
        traceback.print_exc()     


class GridOptionsEditorView(unohelper.Base, XChangesListener):
    def __init__(self, model):
        self.model = model
        self.createChangesListener()
        self.updateView()    
    def updateView(self):
        if self.model is not None:
            print("Grid options editor: data={}".format(self.readModel()))
        else:
            print("Grid options editor: no model set")
    def readModel(self):
        try:
            options = "Option/VisibleGrid", "Resolution/XAxis/Metric", "Resolution/YAxis/Metric", "Subdivision/XAxis", "Subdivision/YAxis"
            values = self.model.getHierarchicalPropertyValues(options)
            return  GridOptions(*values)
        except Exception as e:
            GridOptionsEditor.informUserOfError(e)
            return None 
    def createChangesListener(self):
        self.model.addChangesListener(ChangesListener(self))
class ChangesListener(unohelper.Base, XChangesListener):      
    def __init__(self, cast):
        self.cast = cast                   
    def changesOccurred(self, event):
        print("GridEditor - Listener received changes event containing {} change(s).".format(len(event.Changes)))
        self.cast.updateView()
    def disposing(self, event):
        print("GridEditor - Listener received disposed event: releasing model")
# def changeSomeData(path, cp):
#     try:
#         cu = getUpdatableModel(path, cp)
#         itemnames = cu.getElementNames()
#         for itemname in itemnames:
#             item = cu.getByName(itemname)
#             if isinstance(item, bool):
#                 print("Replacing boolean value: {}".format(itemname))
#                 cu.replaceByName(itemname, False if item else True)
#             elif isinstance(item, int):
#                 item = 9999-item
#                 print("Replacing integer value: {}".format(itemname))
#                 cu.replaceByName(itemname, item)
#         cu.commitChanges()
#         cu.dispose()
#     except:
#         print("Could not change some data in a different view. An exception occurred:")
#         traceback.print_exc() 

    
    
    
# def resetGridConfiguration(cp):
#     path = "/org.openoffice.Office.Calc/Grid"
#     cu = createUpdatableView(path, cp)
#     
#     
#     state = cu.getByHierarchicalName("{}/Option".format(path))
#     state.setPropertyToDefault("VisibleGrid")
#     
#     
#     cu.getByHierarchicalName("{}/Option".format(path)).setPropertyToDefault("VisibleGrid")
#     cu.getByHierarchicalName("Resolution/XAxis").setPropertyToDefault("Metric")
#     cu.getByHierarchicalName("Resolution/YAxis").setPropertyToDefault("Metric")
#     cu.getByHierarchicalName("Subdivision").setAllPropertiesToDefault()
#     cu.commitChanges()
#     cu.dispose()
# def resetGroupExample(cp):
#     try:
#         print("\n--- starting example: reset group data -----------------------------")
#         olddata = readGridConfiguration(cp)
#         resetGridConfiguration(cp)
#         newdata = readGridConfiguration(cp)
#         print("Before reset:   user grid options: {}".format(olddata))
#         print("After reset: default grid options: {}".format(newdata))
#     except:
#         traceback.print_exc()     
    
    
    
    
    

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