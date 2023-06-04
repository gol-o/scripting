' check for initialization
wscript.echo i = empty ' true

' default numeric value autocast to empty
i = 0
wscript.echo i = empty ' true

i = 1
wscript.echo i = empty ' false

' default string value "" autocast to empty
s = ""
wscript.echo s = empty ' true

' object variables
on error resume next
dim t
wscript.echo t is nothing ' no object at this point causes error
set t = createobject( "excel.application" )

t.visible = true
wscript.echo t is nothing ' false, k = nothing not allowed

wscript.stdout.write "shutting down excel..."
wscript.stdin.readline

set t = nothing
wscript.echo t is nothing ' true

wscript.stdout.write "key..."
wscript.stdin.readline

wscript.echo t is nothing ' true

' invalid values
' never compare direct to null (always null)
wscript.echo i = null    ' null
wscript.echo isnull( i ) ' false
i = null
wscript.echo isnull( i ) ' true
