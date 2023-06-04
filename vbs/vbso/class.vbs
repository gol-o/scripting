class Tree
    private nodeId

    ' constructor (arguments not allowed)
    private sub class_initialize( )
        nodeId = 1
    end sub

    ' destructor
    private sub class_terminate( )
         wscript.echo "terminating instance..."
    end sub

    public property let nid( id )
        nodeId = id
    end property

    public sub setId( id )
        if id < 1 or id > 9 then
            wscript.echo "wrong id number range"
        else
            nodeId = id
        end if
    end sub

    public sub out( )
        wscript.echo "Node-Id: ", nodeId
    end sub
end class

dim t
set t = new Tree
t.nid = 33
t.out( )

t.setId( 5 )
t.out( )

t.setId( 10 )
t.out( )
