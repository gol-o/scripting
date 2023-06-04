' start wscript gui (default)
WScript.Echo "test" ' as msgbox
MsgBox "test"

' start cscript commandline scripting host
WScript.Echo "test" ' as text
WScript.StdOut.Write "test"
WScript.StdOut.WriteLine "test"
MsgBox "test"

' input
t = InputBox( "test", "title", "default" )
MsgBox t

' fso object
Set out = CreateObject("Scripting.FileSystemObject").GetStandardStream(1)
out.Write "test"
out.WriteLine "test"

' special character
dim qm
    qm = ""      ' empty
    qm = """"    ' "
    qm = """"""  ' ""
'   qm = """"""" ' error
wscript.echo qm
