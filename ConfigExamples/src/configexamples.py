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

def main(ctx, smgr):
    cp = createProvider(ctx, smgr)
    if checkProvider(cp):
        print("\nStarting examples.")
        readDataExample(cp)
#         browseDataExample(cp)
#         updateGroupExample(cp)
#         resetGroupExample(cp)
        
        
        
        print("\nAll Examples completed.")
    else:
        print("ERROR: Cannot run examples without ConfigurationProvider.")
def createProvider(ctx, smgr):
    return smgr.createInstanceWithContext("com.sun.star.configuration.ConfigurationProvider", ctx)
#     return smgr.createInstanceWithContext("com.sun.star.configuration.DefaultProvider", ctx)
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
def readGridConfiguration(cp):
    ca = createConfigurationView("/org.openoffice.Office.Calc/Grid", cp)
    

    
    visible = ca.getHierarchicalPropertyValue("Option/VisibleGrid")
    

#     reso = ca.getHierarchicalPropertyValue("Resolution")

    reso = ca.getPropertyValue("Resolution")

    reso_elems = reso.getHierarchicalPropertyValues(("XAxis/Metric", "YAxis/Metric"))
    sub = ca.getPropertyValue("Subdivision")
    sub_elems = sub.getPropertyValues(("XAxis", "YAxis"))
    return GridOptions(visible, reso_elems[0], reso_elems[1], sub_elems[0], sub_elems[1])    
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
class IconfigurationProcessor:
    def processValueElement(self, path, values):
        if isinstance(values, tuple):
            print("\tValue: {0} = {{{1}}}".format(path, ", ".join(values)))
        else:
            print("\tValue: {} = {}".format(path, values))
    def processStructuralElement(self, path, elem):
        if hasattr(elem, "getTemplateName") and elem.getTemplateName().endswith("Filter"):
            print("Filter {} ({})".format(elem.getName(), path))
def printRegisteredFilters(cp):
    path = "/org.openoffice.TypeDetection.Filter/Filters"
    browseConfiguration(path, IconfigurationProcessor(), cp)
def browseConfiguration(path, processor, cp):
    ca = createConfigurationView(path, cp)
    browseElementRecursively(ca, processor)
    ca.dispose()
def browseElementRecursively(elem, processor):
    path = elem.getHierarchicalName()
    processor.processStructuralElement(path, elem)
    childnames = elem.getElementNames()
    for childname in childnames:
        child = elem.getByName(childname)
        if hasattr(child, "getTypes"):
            browseElementRecursively(child, processor)
        else:
            childpath = elem.composeHierarchicalName(childname)
            processor.processValueElement(childpath, child)
def updateGroupExample(cp):
    try:
        print("\n--- starting example: update group data --------------")
        editGridOptions(cp)
    except:
        traceback.print_exc()
def createUpdatableView(path, cp):
    node = PropertyValue(Name="nodepath", Value=path)
    return cp.createInstanceWithArguments("com.sun.star.configuration.ConfigurationUpdateAccess", (node,))
class ChangesListener(unohelper.Base, XChangesListener):     
    CANCELED = 0
    SAVE_SETTINGS = 1
    def __init__(self, model):
        self.model = model
        self.updateDisplay()    
    def changesOccurred(self, event):
        print("GridEditor - Listener received changes event containing {} change(s).".format(len(event.Changes)))
        self.updateDisplay()
    def disposing(self, event):
        print("GridEditor - Listener received disposed event: releasing model")
        self.setModel(None)  
    def updateDisplay(self):
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
            self.informUserOfError(e)
            return None 
    def execute(self):
        try:
            print("-- GridEditor executing --")
            self.toggleVisibility()
            print("-- GridEditor done      --")
            return self.SAVE_SETTINGS
        except Exception as e:
            self.informUserOfError(e)
            return self.CANCELED
    def informUserOfError(self, e):
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
        cu = createUpdatableView(path, cp)
        itemnames = cu.getElementNames()
        for itemname in itemnames:
            item = cu.getByName(itemname)
            if isinstance(item, bool):
                print("Replacing integer value: {}".format(itemname))
                cu.replaceByName(itemname, False if item else True)
            elif isinstance(item, int):
                item = 9999-item
                print("Replacing integer value: {}".format(itemname))
                cu.replaceByName(itemname, item)
        cu.commitChanges()
        cu.dispose()
    except:
        print("Could not change some data in a different view. An exception occurred:")
        traceback.print_exc() 
def editGridOptions(cp):
    path = "/org.openoffice.Office.Calc/Grid"
    cu = createUpdatableView(path, cp)
    dialog = ChangesListener(cu)
    cu.addChangesListener(dialog)
    changeSomeData("{}/Subdivision".format(path), cp)
    if dialog.execute()==ChangesListener.SAVE_SETTINGS:
        try:
            cu.commitChanges()
        except Exception as e:
            dialog.informUserOfError(e)        
    cu.removeChangesListener(dialog)    
    cu.dispose()    
def resetGridConfiguration(cp):
    path = "/org.openoffice.Office.Calc/Grid"
    cu = createUpdatableView(path, cp)
    
    
    state = cu.getByHierarchicalName("{}/Option".format(path))
    state.setPropertyToDefault("VisibleGrid")
    
    
    cu.getByHierarchicalName("{}/Option".format(path)).setPropertyToDefault("VisibleGrid")
    cu.getByHierarchicalName("Resolution/XAxis").setPropertyToDefault("Metric")
    cu.getByHierarchicalName("Resolution/YAxis").setPropertyToDefault("Metric")
    cu.getByHierarchicalName("Subdivision").setAllPropertiesToDefault()
    cu.commitChanges()
    cu.dispose()
def resetGroupExample(cp):
    try:
        print("\n--- starting example: reset group data -----------------------------")
        olddata = readGridConfiguration(cp)
        resetGridConfiguration(cp)
        newdata = readGridConfiguration(cp)
        print("Before reset:   user grid options: {}".format(olddata))
        print("After reset: default grid options: {}".format(newdata))
    except:
        traceback.print_exc()     
    
    
    
    
    

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