class node
    public  info
    public  successor
    private data(30)

    private sub class_initialize( )
        fill( )
    end sub

    function fill( )
        for i = 0 to ubound( data )
            data( i ) = chr( 65 + int( 60 * rnd ) )
        next
    end function

    function out( )
        wscript.stdout.write "Data: "
        for each e in data
            wscript.stdout.write e
        next
        wscript.stdout.writeline
        wscript.stdout.writeline "----------------------------------------"
    end function
end class

dim n1
set n1           = new node
n1.info          = "node 1"
n1.successor     = empty

dim n2
set n2           = new node
n2.info          = "node 2"
set n2.successor = n1

dim n3
set n3           = new node
n3.info          = "node 3"
set n3.successor = n2

dim n
set n = n3

while not isempty( n )
    wscript.echo "Nodeinfo:", n.info
    n.out( )
    set n = n.successor
wend
