import win32com.client
import pythoncom
import win32

class CadUtils:
    def __init__(self, app_name):
        self.wincad = win32com.client.Dispatch(app_name)
        self.doc = self.wincad.ActiveDocument
        self.msp = self.doc.ModelSpace
        self.wcs = True

    def vtpnt(self, x, y, z=0):
        """创建点对象"""
        return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x, y, z))

    def cad_ucs(self):
        """获取UCS信息"""
        # ... (原cad_ucs方法的代码) 

    def vtobj(self, obj):
        """转化为对象数组"""
        return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, obj)

    def autocad(self):
        """获取CAD应用程序实例"""
        try:
            cad = win32com.client.GetActiveObject("AutoCAD.Application")
            cad.Visible = True
        except:
            cad = win32com.client.Dispatch("AutoCAD.Application")
            cad.Visible = True
        return cad 