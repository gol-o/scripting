dim x
x = 1

' create global sub on the fly
executeglobal "sub s( ): wscript.echo ""x ="", x: end sub"
s( )

sub t( )
    ' procedure still available in global namespace
    executeglobal "sub u( ): wscript.echo ""x ="", x: end sub"
    u( )
end sub
t( )
u( ) ' globally available

' locally created sub in the context of v
sub v( )
    execute "sub w( ): wscript.echo ""x ="", x: end sub"
    w( )
end sub
v( )

on error resume next
w( ) ' error, not available in global namespace
if err.raise then wscript.echo "runtime error"
on error goto 0
