set shell = wscript.createobject( "wscript.shell" )
set shell =         createobject( "wscript.shell" ) ' short form

' get username
user = shell.expandenvironmentstrings( "%username%" )
wscript.echo user

' get desktop path
desktop = shell.specialfolders( "desktop" )
wscript.echo desktop

' execute shell command
set exec = shell.exec( "cmd.exe /c dir" )
do
    wscript.echo exec.stdout.readline( )
loop while not exec.stdout.atendofstream
