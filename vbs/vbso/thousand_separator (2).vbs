'================================================================================
' number -> number_leading_part for number < 1000
'        -> f( number_leading_part ) . number_rest_part
' 
' number_leading_part = part by integer division
' number_rest_part    = 3 digits modulus
'================================================================================

dim n

sub splitter( number )
    if number < 1000 then
        wscript.stdout.write number
    else
        n = n + 1
        wscript.echo "Partition", n & ":", number, number \ 1000, number mod 1000

        splitter( number \ 1000 )
        wscript.stdout.write "." & number mod 1000
    end if
end sub

splitter( 1234568900 )
wscript.echo
