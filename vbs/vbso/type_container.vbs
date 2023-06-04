' reference: Microsoft Scripting Runtime (\system32\scrrun.dll)

option explicit

function iif( expression, truepart, falsepart )
    if expression then
        iif = truepart
    else
        iif = falsepart
    end if
end function

' define a type structure
dim types
set types = createobject( "scripting.dictionary" )
types.add "t_int"    , 0
types.add "t_date"   , 1
types.add "t_string" , 2
    
' read single character
dim ch
wscript.echo "input single character"
ch = wscript.stdin.read( 1 ) ' discard remaining characters
wscript.stdin.readline

' try to convert
dim dt_type
on error resume next

ch = cint( ch )
if err.number > 0 then
    err.clear
else
    if isempty( dt_type ) then
        dt_type = types.item( "t_int" )
    end if
end if

if isdate( ch ) then ' cdate( ) doesn't raise an error
    if isempty( dt_type ) then
        dt_type = types.item( "t_date" )
    end if
end if

' ...

if not isnumeric( ch ) then
    if isempty( dt_type ) then
        dt_type = types.item( "t_string" )
    end if
end if

select case dt_type
    case types.item( "t_int" )
        wscript.echo "your input was of numeric type: " & ch
    case types.item( "t_date" )
        wscript.echo "your input was of date type: " & ch
    case types.item( "t_string" )
        wscript.echo "your input was of text type: " & ch
end select

' different approach regex
on error goto 0
dim regex
set regex = new regexp
with regex
    .pattern = "...[0-9]{3}"
    .global = true
end with
wscript.echo iif( regex.test( "aaa123" ), "matches", "no match" )
wscript.echo iif( regex.test( "aaaa23" ), "matches", "no match" )

' digit pattern
regex.pattern = "\d+$" ' any number
wscript.echo iif( regex.test( "1234" ), "matches", "no match" )
wscript.echo iif( regex.test( "1234a" ), "matches", "no match" )
