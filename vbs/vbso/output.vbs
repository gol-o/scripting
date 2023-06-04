' start wscript gui (default)
wscript.echo "test" ' as msgbox
msgbox "test"

' start cscript commandline scripting host
wscript.echo "test" ' as text
wscript.stdout.write "test"
wscript.stdout.writeline "test"
msgbox "test"

' input
t = inputbox( "test", "title", "default" )
msgbox t

' fso object
set out = createobject("scripting.filesystemobject").getstandardstream(1)
out.write "test"
out.writeline "test"

' special character
dim qm
    qm = ""      ' empty
    qm = """"    ' "
    qm = """"""  ' ""
'   qm = """"""" ' error
wscript.echo qm
