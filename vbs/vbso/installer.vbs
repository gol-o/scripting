'Skript-Datei nicht in UTF-8 Kodierung speichern

'Start
Call Transfer()

'global Const
Const host = "http://sourcefiles.ifastnet.com/"
Const StdPath = "C:\Dokumente und Einstellungen\Admin\Desktop\"

'user-defined IIf( ) function
Function IIf(expr, truepart, falsepart)
    IIf = falsepart
    If expr Then IIf = truepart
End Function

'user-defined Val( ) function
Function Val(Src)
    With New RegExp
        .IgnoreCase = True
        .Pattern = "^[+-]?\d+(\.\d*)?([DE][+-]?\d+)?"
        With .Execute(Src)
            If .Count = 1 Then Val = Eval(.Item(0)) Else Val = "Error!"
        End With
    End With
End Function

'------------------------------------------------------
' Filetransfer
'------------------------------------------------------
Sub Transfer()
    Dim Folder, SubFolder, CourseList, CourseNr
    CourseList = "Access:1" & vbCrLf & _
                 "Excel:2" & vbCrLf & _
                 "Powerpoint:3" & vbCrLf & _
                 "Word:4" & vbCrLf & _
                 "C:5" & vbCrLf & _
                 "C++:6" & vbCrLf & _
                 "C#:7" & vbCrLf & _
                 "Java:8" & vbCrLf & _
                            "VB:9" & vbCrLf & _
                            "VB.NET:10" & vbCrLf & _
                            "JavaScript:11" & vbCrLf & _
                            "VBScript:12" & vbCrLf & _
                            "PHP:13" & vbCrLf & _
                            "Erase Folder:14/KursNr"
    CourseNr = InputBox(CourseList, "OfflineInstaller 1.0.2 - Kurswahl", 1)
    Folder = InputBox("Zielpfad", "Please standby while loading...", "C:\Dokumente und Einstellungen\Admin\Desktop\")
    Folder = Folder & IIf(Right(Folder, 1) <> "\", "\", "") 'falls "\" fehlt, anfügen

    Dim CourseID, Course
    If Val(CourseNr) = 14 Then
        CourseID = Split(CourseList, vbCrLf)(Val(Mid(CourseNr, 4)) - 1)
        Course = Left(CourseID, InStr(CourseID, ":") - 1) & "KURS"
        EraseFolder Folder, Course
        Exit Sub
    End If
    CourseID = Split(CourseList, vbCrLf)(CourseNr - 1)
    Course = Left(CourseID, InStr(CourseID, ":") - 1) & "KURS\"

    Select Case CourseNr
        Case 1: SubFolder = "acc"
        Case 2: SubFolder = "xl"
        Case 3: SubFolder = "ppt"
        Case 4: SubFolder = "wrd"
        Case 5: SubFolder = "c"
        Case 6: SubFolder = "cpp"
        Case 7: SubFolder = "cs"
        Case 8: SubFolder = "jav"
        Case 9: SubFolder = "vb"
        Case 10: SubFolder = "vbn"
        Case 11: SubFolder = "js"
        Case 12: SubFolder = "vbs"
        Case 13: SubFolder = "php"
    End Select
    SubFolder = SubFolder & "sources"

    'Dozentenordner erstellen
    Call MakeFolder(Folder & "DOZENT")

    'Kursordner erstellen
    Folder = Folder & UCase(Course) 'Ordner in Grossbuchstaben
    Call MakeFolder(Folder)
    Call MakeFolder(Folder & "Coursefiles\")
    Call MakeFolder(Folder & "Coursefiles\" & SubFolder)

    'Imageordner erstellen
    Call MakeFolder(Folder & "Images")

    Dim File
    'Verzeichnis laden
    For Each File In ReadWebDirectory(host & "Coursefiles/" & SubFolder)
        'File downloaden
        Call Download(File, host & "Coursefiles/" & SubFolder & "/", Folder & "Coursefiles\" & SubFolder & "\")
    Next

    'Restliche Dateien  
    Call Download( "index.html", host, Folder)
    Call Download ("NoScript.htm",  host, Folder)
    Call Download ("schedule.htm",  host & "Coursefiles/", Folder & "Coursefiles\")
    Call Download ("Referenzen.jpg", host & "Images/", Folder & "Images\")
    Call Download ("guru.jpg", host & "Images/", Folder & "Images\")

    MsgBox "Complete!"
End Sub

'------------------------------------------------------
' Verzeichnis erstellen
'------------------------------------------------------
Sub MakeFolder(f_name)
    On Error Resume Next
    Err.Clear
    Dim fso, fldr
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fldr = fso.CreateFolder(f_name)
    If Err.Number <> 0 Then MsgBox ("Cannot create Folder!")
End Sub

'------------------------------------------------------
' Verzeichnis/Unterverzeichnis(se) löschen
'------------------------------------------------------
Sub EraseFolder(f_name, c_name)
    On Error Resume Next
    Err.Clear
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.DeleteFolder (f_name & c_name)
    fso.DeleteFolder (f_name & "DOZENT")
    If Err.Number <> 0 Then MsgBox ("Cannot delete Folder!")
End Sub

'------------------------------------------------------
' Verzeichnis auf dem Server holen
' und aufbereiten
'------------------------------------------------------
Function ReadWebDirectory(SourcePath)
    On Error Resume Next
    Err.Clear
    Dim oXML, strResponse
    Set oXML = CreateObject("MSXML2.XMLHTTP") 'an ActiveX component that performs HTTP requests
    With oXML
        .Open "GET", SourcePath, False
        .Send
        strResponse = .ResponseText  'returns the body of the response as an array
        If Err.Number <> 0 Or .Status <> 200 Then Exit Function
    End With

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Filenamen extrahieren
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set oRegExp = New RegExp
    oRegExp.Pattern = "href\s*=\s*""[^""?\/=]+"""
    oRegExp.IgnoreCase = True
    oRegExp.Global = True
    Set Matches = oRegExp.Execute(strResponse)
    Dim DirList(), i
    i = 0
    ReDim DirList(Matches.Count - 1)
    For Each Match In Matches
        'Array aufbauen, Leerzeichen, href und " entfernen
        DirList(i) = Replace(Replace(Replace(Match.Value, " ", ""), "href=", ""), Chr(34), "")
        i = i + 1
    Next
    ReadWebDirectory = DirList
End Function

'------------------------------------------------------
' HTTP Download
'------------------------------------------------------
Sub Download(File, SourcePath, DestPath)
    Dim Response
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Read Data
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Err.Clear
    With CreateObject("MSXML2.XMLHTTP")
        .Open "GET", SourcePath & File, False
        .Send
        Response = .ResponseBody 'returns the body of the response as an array
        If Err.Number <> 0 Or .Status <> 200 Then MsgBox ("Error reading Data!"): Exit Sub 
    End With

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Write Data
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Const adTypeBinary = 1
    Const adSaveCreateOverwrite = 2
    Const adModeReadWrite = 3

    Set oStream = CreateObject("ADODB.Stream")
    Err.Clear
    With oStream
        .Type = adTypeBinary
        .Mode = adModeReadWrite
        .Open
        .Write Response
        .SaveToFile DestPath & File, adSaveCreateOverwrite
        .Close
        If Err.Number <> 0 Then MsgBox ("Error writing Data!"): Exit Sub 
    End With
End Sub
