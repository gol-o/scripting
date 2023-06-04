function power( digit, exponent )
    power = digit ^ exponent
end function

sub getdate( )
    wscript.echo date
end sub

' function/sub references
set ref = getref( "getdate" )
ref( )

set ref = getref( "power" )
wscript.echo ref( 4, 3 )

' variant/object type
vtype = 0
wscript.echo typename( vtype )
wscript.echo vartype ( vtype )

vtype = array( )
wscript.echo typename( vtype )
wscript.echo vartype ( vtype ) ' vartype( array ) = vbArray = 8192 + vartype( stored type )
wscript.echo vbArray + vbVariant ' 8204

' set
' needed for object type
' vtype = getref( "power" ) → error
wscript.echo typename( getref( "power" ) ) ' vbObject
wscript.echo vartype ( getref( "power" ) ) ' 9 
wscript.echo vbObject ' 9

vtype = array( )
' set vtype = array( ) → error, object type needed

wscript.echo typename( vtype ) ' vbArray, Variant()
wscript.echo vartype ( vtype ) ' 8192
