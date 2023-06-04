employee = inputbox( "Name and Department" )

' check for delimiter
pos = instr( employee, ";" )
if not cbool( pos ) then ' try different delimiter
    pos = instr( employee, "," )
    if not cbool( pos ) then
        pos = instr( employee, " " )
        if not cbool( pos ) then
            msgbox "delimiter not included"
        end if
    end if
end if

filename = left( employee, pos - 1 )
desktop  = createobject( "wscript.shell" ).specialfolders( "desktop" )
set file = createobject( "scripting.filesystemobject" ). _
           createtextfile( desktop + "/" + filename )

file.writeline( "employee record, timestamp: " + formatdatetime( now, vblongtime ) )
file.writeline( "---------------------------------------" )
file.writeline( "name: " + filename ) 
file.writeline( "dep : " + mid( employee, pos + 1 ) ) 
file.close( )

msgbox "record written"
