strIn = "aaa111bbbb2222ggggg1111222aaa"

set dict = createobject( "scripting.dictionary" )

for i = 1 to len( strIn )
    ch = mid( strIn, i, 1 )

    if dict.exists( cstr( asc( ch ) ) ) then
        ' key has already been inserted
        ' therefore counter need to be incremented
         dict.item( asc( ch ) ) = dict.item( asc( ch ) ) + 1
    else
        ' insert new key
        dict.add cstr( asc( ch ) ), ch
    end if
next

' pairwise output
' wscript.echo dict.count
keys = dict.keys

for i = 1 to dict.count
'     wscript.echo dict.item( i )
next

for each k in keys
'     wscript.echo k 
next

for i = 1 to 10 'dict.count
    wscript.echo dict.item(i) 
next

wscript.quit


for each e in dict
    wscript.echo e
next
wscript.quit
for each key in dict.keys
    wscript.echo key, dict.item( key )
next

