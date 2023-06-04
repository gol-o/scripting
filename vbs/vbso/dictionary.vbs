' random guid
function guid( )
    ' create guid, remove {, }, -, limit size
    guid = replace( mid( createobject( "scriptlet.typelib" ).guid, 2, 24 ), "-", "" )
End Function

'================================================================================
' Dictionary
' • key/value associative container
' • key: any_type but array
' • val: any_type
'
' Properties
' • comparing:       comparemode = vbtextcompare, vbbinarycompare, ...
'                    • comparison of string keys
' • set key:         key( key: any_type ): key_type
'                    • read/write key,
'                    • throws exception if key doesn't exist
' • access by key:   item( key: any_type ): value_type
'                    • read/write associated value,
'                    • adds entry if key hasn't been used
' • counting:        count
'
' Methods
' • check existance: exists( key: any_type ): boolean
' • insert:          add key: any_type, value: value_type
'                    • throws exception if key has already been used
' • delete entry:    remove( key: any_type ), removeall
' • array of keys:   keys( ): variant( )
' • array of values: items( ): variant( )
'================================================================================

set dic = createobject( "scripting.dictionary" )

' comparing keys
dic.comparemode = vbbinarycompare ' default setting
dic.add "x", 1
dic.add "X", 2 ' different key, "x" <> "X"
wscript.echo dic( "x" ), dic( "X" )

' before changing mode, dictionary need to be cleared to
' avoid data corruption on existing data (based on binary comparing)
dic.removeall

dic.comparemode = vbtextcompare
on error resume next
dic.add "y", 1
dic.add "Y", 2
if err.number > 0 then
    wscript.echo "error, key already used"
    err.clear
end if
on error goto 0
wscript.echo dic( "y" ), dic( "Y" ) ' treated as one key

' improves flexibility
dic.add "Smith", "4493 Collins Avenue"
wscript.echo dic( "SMITH" ), ",",  dic( "smith" )

' different key types
dic.add "1", "a"
dic.add "2", "b"
dic.add  1 , "c"
dic.add  2 , "d"

on error resume next
dic.add  2 , "e"
if err.number > 0 then
    wscript.echo "error, key already used"
    err.clear
end if
on error goto 0

' keys
wscript.echo dic.keys( )( 0 ) ' extract first key

for each k in dic.keys
    wscript.stdout.write k
    wscript.stdout.write " "
next

wscript.echo

' values (items)
for each i in dic.items
    wscript.stdout.write i
    wscript.stdout.write " "
next

wscript.echo

' access through key
for each e in dic
    ' show key
    wscript.stdout.write e
    wscript.stdout.write " "

    ' show value
    ' dict_var.item( key ) [item method]
    ' dict_var     ( key ) [short form ]
    wscript.echo dic.item( e )
next

' different value types
dic.add 3, #12/01/2000#

for each e in dic
    wscript.stdout.write "key: " & _ 
        e & _
        ", value: " & _
        dic( e ) & _
        ", value type: " & _
        typename( dic( e ) )
    wscript.echo
next

dic.add 4, array( 11, 13, 17 ) ' would cause an error, needs specific output

' adapted output
' select case throws an exception if selector is of string
' type compared with first case numeric type → rearrange cases
' or perform comparison on string type
for each e in dic
    wscript.stdout.write e & " "

    select case e
        case "3" ' date
            wscript.stdout.write formatdatetime( dic( e ), vblongdate )
        case "4" ' array
            wscript.stdout.write dic( e )( 0 )
        case else
            wscript.stdout.write dic( e )
    end select

    wscript.echo
next

dic.removeall

class Customer
    private m_id
    private m_company

    ' constructor
    private sub class_initialize( )
        p_ctor 0, empty 
    end sub

    ' pseudo constructor
    public function p_ctor( new_id, new_company )
        m_id       = new_id
        m_company  = new_company
        set p_ctor = me
    end function

    public property get id( )
        id = m_id 
    end property

    public property let id( new_id )
        m_id = new_id
    end property

    public property get company( )
        company = m_company
    end property

    public property let company( new_company )
        m_company = new_company
    end property

    public sub tostring( )
        ws = space( 20 - len( company ) )
        wscript.echo "  +----------------------------------+"
        wscript.echo "  |  CUSTOMER PROFILE                |"
        wscript.echo "  |                                  |"
        wscript.echo "  |  Id       ", id, space( 16 ),   "|"
        wscript.echo "  |  Company  ", company, ws,       "|"
        wscript.echo "  +----------------------------------+"
    end sub
end class

set customers = createobject( "scripting.dictionary" )
set items = createobject( "scripting.dictionary" )

set tspoon = new Customer
tspoon.id = 7733
tspoon.company = "Teaspoon Ltd."

set blue = new Customer
blue.id = 9911
blue.company = "BlueSight Beverages"

customers.add tspoon.id, tspoon
customers.add blue.id  , blue

wscript.echo vblf & "Total #Customers:", customers.count
for each e in customers
    wscript.echo
    wscript.echo "Key: ", e
    customers( e ).tostring( )
next

randomize
id = 1000
for i = 1 to 10
    id = id + int( rnd * 500 ) + 1
    customers.add id, (new Customer).p_ctor( id, guid( ) )
next

wscript.echo vblf & "Total #Customers:", customers.count
for each e in customers
    wscript.echo
    wscript.echo "Key: ", e
    customers( e ).tostring( )
next

' nested structure
set a = createobject( "scripting.dictionary" )
set b = createobject( "scripting.dictionary" )
set c = createobject( "scripting.dictionary" )
set d = createobject( "scripting.dictionary" )
set e = createobject( "scripting.dictionary" )

sub layertraverse( dic, ws )
    if typename( dic ) <> "Dictionary" then ' default binary comparison
        wscript.stdout.write dic & " "
        exit sub
    end if

    ' ws: whitespace indentation
    if ws > 0 then wscript.echo
    wscript.stdout.write space( ws ) + "("
    wscript.echo
    ws = ws + 2
    wscript.stdout.write space( ws )

    for each e in dic
        wscript.stdout.write e & ":"
        layertraverse dic( e ), ws + 2
    next

    wscript.echo
    ws = ws - 2
    wscript.stdout.write space( ws ) + ")"
    wscript.echo
    wscript.stdout.write space( ws )
end sub

a.add 1, "a1"
a.add 2, "a2"
a.add 3, b
a.add 4, c
a.add 5, "a5"
a.add 6, d
a.add 7, "a7"
b.add 1, "b1"
b.add 2, "b2"
b.add 3, "b3"
c.add 1, "c1"
c.add 2, "c2"
d.add 1, "d1"
d.add 2, e
d.add 3, "d3"
e.add 1, "e1"
e.add 2, "e2"
e.add 3, "e3"

layertraverse a, 0
wscript.echo
