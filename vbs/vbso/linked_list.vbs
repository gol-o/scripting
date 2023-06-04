p=empty

set p = nothing
set p=empty


wscript.quit
on error resume next
randomize

'================================================================================
' Node structure
' • constructor
' • fill object with random data
' • stringifier
'================================================================================
class node
    public  info
    public  successor
    private data( 30 )

    private sub class_initialize( )
        fill( )
    end sub

    public function fill( )
        for i = 0 to ubound( data )
            data( i ) = chr( 65 + int( 60 * rnd ) )
        next
    end function

    public function out( )
        wscript.stdout.write "Data: "
        for each e in data
            wscript.stdout.write e
        next
        wscript.stdout.writeline
        wscript.stdout.writeline "----------------------------------------"
    end function
end class

'================================================================================
' set up linked list
'================================================================================
dim n1
set n1           = new node
n1.info          = "node 1"
set n1.successor = nothing

dim n2
set n2           = new node
n2.info          = "node 2"
set n2.successor = n1

dim n3
set n3           = new node
n3.info          = "node 3"
set n3.successor = n2

'================================================================================
' loop version with check for emptiness
' object variable raises error when assigned to empty
'================================================================================
dim n
set n = n3
do while not isempty( n )
    wscript.echo "Nodeinfo:", n.info
    n.out( )
    set n = n.successor
    if err.number then wscript.stdout.writeline "caught improper assignment... stop": exit do
loop

'================================================================================
' enhanced version with 'nothing' doesn't require error handling 
'================================================================================
wscript.echo
set n = n3
do while not n is nothing
    wscript.echo "Nodeinfo:", n.info
    n.out( )
    set n = n.successor
loop

'================================================================================
' recursively
'================================================================================
sub traverseList( n )
    if n is nothing then exit sub
    traverseList( n.successor )
    wscript.echo "Nodeinfo:", n.info
    n.out( )
end sub

wscript.echo
set n = n3
traverseList( n )
