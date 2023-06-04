function iif( boolexp, truepart, falsepart )
    if boolexp then iif = truepart else iif = falsepart
end function

class Tree
    private nodeId

    private sub class_initialize( )
        nodeId = 1
    end sub

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

t.setId( 10 )
t.out( )

wscript.echo getuilanguage( )
wscript.echo getlocale( )

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Scripting.FileSystemObject
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
dim fso
set fso = wscript.createobject( "scripting.filesystemobject" )

' folder
wscript.echo fso.getparentfoldername( wscript.scriptfullname )
wscript.echo fso.getbasename( wscript.scriptname )

' file create 
dim file
dim path

path = fso.getparentfoldername( wscript.scriptfullname ) ' full path to the script currently being used
wscript.echo "current folderpath:", path

dim filename
filename = path & "\data.txt" 
wscript.echo "target filename:", filename
set file = fso.createtextfile( filename, true )
file.writeline( "123, nike, detroit" )
file.close( )

' read entiry file
wscript.echo fso.opentextfile( filename ).readall( )

wscript.quit


path = fso.getparentfoldername( wscript.scriptfullname )
wscript.echo path
wscript.echo "file", iif( fso.fileexists( path ), "exists", "not exists" )
dim stream
const write = 1
set stream = fso.opentextfile( path , write, false, -1)
stream.write( 123 )
stream.close( )

const readi = 2
set stream = fso.opentextfile( path , readi, false, -1)
while not stream.atendofstream
    wscript.echo stream.readline
wend
