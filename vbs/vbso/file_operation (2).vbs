'================================================================================
' Microsoft Scripting Runtime Library
' File: Scrrun.dll
'
' Objects
' • Drive
' • Folder
' • File
' • Drive
' • TextStream
'================================================================================

' property scriptfullname: script name and storage path
' script storage location
location = wscript.scriptfullname ' vartype string
wscript.echo "script location:", location

' property scriptname: script name
name = wscript.scriptname ' vartype string
wscript.echo "script name:", name

' getabsolutepathname
' path and filename
with createobject( "scripting.filesystemobject" )
    wscript.echo .getabsolutepathname( wscript.scriptfullname )
    wscript.echo .getabsolutepathname( ".\" ) ' current folder absolute path
    wscript.echo .getabsolutepathname( "..\..\" )
end with

' getparentfoldername
with createobject( "scripting.filesystemobject" )
    wscript.echo .getparentfoldername( "..\..\" ) ' ..
    wscript.echo .getparentfoldername( wscript.scriptfullname) ' folder where file located in
end with

' getbasename
' input: folder-structure → deepest level
'        complete path    → filename (without extension)
' doesn't validate folder/file existence
with createobject( "scripting.filesystemobject" )
    wscript.echo .getbasename( "..\..\"      ) ' .
    wscript.echo .getbasename( "C:\Users\"   ) ' Users
    wscript.echo .getbasename( "C:\data.txt" ) ' data
    wscript.echo .getbasename( "XX:\YY\"     ) ' YY
end with

' create file
with createobject( "scripting.filesystemobject" )
    path = .getparentfoldername( wscript.scriptfullname ) & "\data.txt" 
    set file = .createtextfile( path, true )
end with
file.writeline( "123, nike, detroit" )
file.close( )

' read entiry file
content = createobject( "scripting.filesystemobject" ).opentextfile( path ).readall( )
wscript.echo content

function iif( boolexp, truepart, falsepart )
    if boolexp then iif = truepart else iif = falsepart
end function

' check file existence
set fso = createobject( "scripting.filesystemobject" )
wscript.echo "file", iif( fso.fileexists( path ), "exists", "not exists" )

' write/read stream
const write = 2
set stream = fso.opentextfile( path, write, false, 0 )
stream.writeline( 123 )
stream.writeline( 456 )
stream.writeline( "abc" & 123 )
stream.close( )

const read = 1
set stream = fso.opentextfile( path, read, false, 0 )
while not stream.atendofstream
    wscript.echo stream.readline
wend
stream.close( )
