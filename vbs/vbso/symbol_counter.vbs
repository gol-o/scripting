strIn = "aaa111bbbb2222ggggg1111222aaa"

set dict = createobject( "scripting.dictionary" )

for i = 1 to len( strIn )
    ch = mid( strIn, i, 1 )

    if dict.exists( ch ) then
        ' key has already been inserted
        ' therefore counter needs to be incremented
        dict( ch ) = dict( ch ) + 1
    else
        ' insert new key
        dict.add ch, 1
    end if
next

for each e in dict
    wscript.echo e & ":", dict( e )
next
