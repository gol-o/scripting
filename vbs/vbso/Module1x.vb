'-------------------------------------------------------------------------------
' Tool Chain
' Code.vbs → Preprocessor → Preprocessed → Compiler → MSIL(.dll, .exe)
'          → CLR/JITC → Execution 
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
' Command Line
'-------------------------------------------------------------------------------
' 1. developer command prompt
' 2. cd %homepath%/desktop
' 3. notepad message.vb
' 4. vbc message.vb 
' 5. message.exe
' 6. ildasm message.exe /html /out=message.htm
#If 0 Then
Module Message
    Sub Main()
        MsgBox(Now)
    End Sub
End Module
#End If

'-------------------------------------------------------------------------------
' Command Line Arguments
'-------------------------------------------------------------------------------
#If 0 Then
Module Message
    Sub Main(args As String())
        MsgBox("Hello " + args(0) + vbCr + "Today is: " + Now)
    End Sub
End Module
#End If

'-------------------------------------------------------------------------------
' Windows Scripting Host
'-------------------------------------------------------------------------------
#If 0 Then
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
set file = createobject( "scripting.filesystemobject" ).createtextfile( desktop + "/" + filename )

file.writeline( "employee record, timestamp: " + formatdatetime( now, vblongtime ) )
file.writeline( "---------------------------------------" )
file.writeline( "name: " + filename ) 
file.writeline( "dep : " + mid( employee, pos + 1 ) ) 
file.close( )

msgbox "record written"
#End If

'-------------------------------------------------------------------------------
' Framework System Library
' must precede any declarations
'-------------------------------------------------------------------------------
Imports System.Console
Imports System.IO.IsolatedStorage
Imports System.Runtime.InteropServices

' Project Settings: Startobject 
#Const STARTMODULE = False
#Const STARTCLASS = False
#Const STARTCOMMAND = False
#Const STARTPROJECT = False

'-------------------------------------------------------------------------------
' Startmodule
'-------------------------------------------------------------------------------
Module MainModule
#If STARTMODULE Then
    Sub Main()
        WriteLine("Module Main")

        MsgBox("Hello " + Environment.UserName + vbCr +
               "It's " + DateTime.Now.ToShortTimeString)

        If Hour(Now) < 20 Then
            MsgBox("System shutdown in " + (#20:00# - TimeOfDay).ToString() + " hours")
        End If

        ' VB
        WriteLine(Global.Microsoft.VisualBasic.Hour(Global.Microsoft.VisualBasic.DateAndTime.Now))

        ' .NET Framework
        WriteLine(Global.System.DateTime.Now.Hour)
    End Sub
#End If
End Module

'-------------------------------------------------------------------------------
' Startclass
'-------------------------------------------------------------------------------
Class Application
#If STARTCLASS Then
    Shared Sub Main()
        Console.WriteLine("Class Main")
    End Sub
#End If
End Class

'-------------------------------------------------------------------------------
' IDE Command Line Arguments
'
' Project Properties/Debug/Command line arguments
' a) start from command line
' b) solution explorer/folder view/app_folder/bin/debug/project_name.exe
'    open with 'windows command processor'
'-------------------------------------------------------------------------------
Module CommandLineArgs
#If STARTCOMMAND Then
    Sub Main(args As String())
        ' direct access
        WriteLine("First argument: " & args(0))

        ' loop
        For Each arg In args
            WriteLine(arg)
        Next
    End Sub
#End If
End Module

'-------------------------------------------------------------------------------
' Preventing name clashes
'-------------------------------------------------------------------------------
Namespace Development
    Module Project
        Sub StartApp()
            WriteLine("Dev Dep")
        End Sub
    End Module
End Namespace

Namespace CustomerSpace
    Module Project
        Public Sub StartApp()
            WriteLine("Customer Dep")
        End Sub
    End Module
End Namespace

Module ProjectAdmin
#If STARTPROJECT Then
    Sub Main()
        ' calling
        Call Development.Project.StartApp()
        ' or
        Call Development.StartApp()
        ' or
        Development.StartApp()

        CustomerSpace.StartApp()
    End Sub
#End If
End Module

'-------------------------------------------------------------------------------
' Set/Determine Datatype
' • VarType()
' • Typename()
' • GetType()
' • TypeOf Operator
'-------------------------------------------------------------------------------
Module Datatypes
#If 0 Then
    Sub Main()
        Dim var ' Object type, not initialized

        ' VB
        Dim vt As Microsoft.VisualBasic.VariantType ' Enumeration
        vt = Microsoft.VisualBasic.Information.VarType(var)
        WriteLine(vt)            ' 9
        WriteLine(vt.ToString()) ' Object, optimized with toString()

        Dim vtName As String
        vtName = Microsoft.VisualBasic.Information.TypeName(var)
        WriteLine(vtName)        ' Nothing

        ' .NET
        Dim st As System.Type
        Try
            st = var.GetType()   ' works only if instantiated
        Catch
        End Try

        ' TypeOf should be avoided for uninitialized objects
        Dim checkType As Boolean = TypeOf var IsNot Integer
        WriteLine(checkType)

        checkType = TypeOf var Is Object
        WriteLine("Is" + IIf(checkType, " ", " not ") + "of type Object") ' False

        var = New Object
        checkType = TypeOf var Is Object
        WriteLine("Is" + IIf(checkType, " ", " not ") + "of type Object") ' True

        var = 0

        WriteLine(VarType(var))          ' 3
        WriteLine(TypeName(var))         ' Integer
        WriteLine(var.GetType())         ' System.Int32
        WriteLine(TypeOf var Is Integer) ' True

        ' Short forms
        Dim a  ' Object
        Dim b% ' Integer
        Dim c! ' Single
        Dim d$ ' String
        Dim e& ' Long
        Dim f# ' Double

        ' Default initialization
        ' Set Nothing ↔ set default value
        Dim i As Integer = Nothing
        WriteLine(i) ' 0

        Dim location$ = Nothing
        WriteLine(location = String.Empty) ' True
        WriteLine(location = "")           ' True
        Try
            WriteLine(location.Length)     ' Error, throws exception
        Catch
        End Try
        WriteLine(Len(location))           ' 0
    End Sub
#End If
End Module

'-------------------------------------------------------------------------------
' Visibility/Linkage/Scope
'-------------------------------------------------------------------------------

' Module level
Module CompanyData
    ' public
    Public company

    ' private
    Dim hotline

    ' private
    Private location
End Module

' Class level
Class Network
    Public ip
    Dim subnet
End Class

Module BusinessApp
#If 0 Then
    Sub Main()
        CompanyData.company = "Milwaukee, Blue Velvet Beverages"
        company = "Milwaukee, Blue Velvet Beverages"

        hotline = "+1 414-334-552" ' Error, private
        location = "823 East Hamilton Street, Milwaukee, WI 53202, United States" ' Error, private
        ip = "200.11.200.3" ' Error, not declared

        Dim net As New Network
        net.ip = "200.11.200.3"
        WriteLine(net.ip)

        net.subnet = "255.255.0.0" ' error
    End Sub
#End If
End Module

' Procedure level
Module Shipping
#If 0 Then
    Const containerType$ = "HighCube"
    Public deliveryFee@ = 70.44

    Sub Supplier()
        Const shippingLocation$ = "Amsterdam" ' local constant
    End Sub

    Sub CustomerSetUp()
        Dim firstLocation = shippingLocation ' Error, not visible
        Dim container = containerType
    End Sub

    ' Decision between local and global definiton
    Sub SetFee()
        Dim deliveryFee@ = 50.33

        WriteLine(deliveryFee)
        WriteLine(Shipping.deliveryFee)
    End Sub

    Sub Main()
        Call SetFee()
    End Sub
#End If
End Module

'-------------------------------------------------------------------------------
' Array
'-------------------------------------------------------------------------------
Module Array
#If 0 Then
    Sub Main()
        Dim a(0 To 3) ' array of objects, always starts at 0
        a(0) = New Integer
        a(1) = "test"
        a(2) = 0.0001D
        a(3) = True

        WriteLine(a.GetType())
        WriteLine(a(0).GetType())
        WriteLine(a(1).GetType())
        WriteLine(a(2).GetType())
        WriteLine(a(3).GetType())

        ' alternatively
        Dim b() = New Integer(2) {1, 2, 3} ' only without explicit bounds
        ' or
        Dim c = New Integer() {4, 5}

        ' get lower/upper bounds
        WriteLine(b.GetLowerBound(0))
        WriteLine(b.GetUpperBound(0))

        ' get number of elements
        WriteLine(b.Length)
        WriteLine(b.Count)
        WriteLine(b.GetLength(0)) ' for specific dimension

        ' copy references
        WriteLine(c(0))
        c = b
        b(0) = 0
        WriteLine(c(0))

        Dim list = {1, 2, 3, 4, 5}

        For i = 1 To list.Length
            Write(i, " ")
        Next
        WriteLine()

        For Each item As String In list
            Write(item, " ")
        Next
        WriteLine()

        ' Join
        WriteLine(String.Join("-", list))

        ' List sequencer
        list.ToList().ForEach(Sub(e) Write(e & ", "))
        WriteLine()

        list.ToList().ForEach(Sub(e) Write(e & IIf(e < 5, ", ", "")))
        WriteLine()

        ' Redimension
        Dim metals() = {"Gold", "Silver", "Platinum", "Palladium"}

        Try
            metals(metals.Length) = "Ruthenium"
        Catch ex As Exception
            WriteLine(ex.Message)
        End Try

        ReDim Preserve metals(metals.Length)
        metals(4) = "Ruthenium"
        WriteLine(String.Join("-", metals))

        ' Returning locally allocated array
        Dim licences
        licences = GetNewLicenceKeys()
        For index = 1 To GetNewLicenceKeys.GetUpperBound(0)
            WriteLine($"{index:D4}: {licences(index)}")
        Next

        ' Multidimensional rectangular arrays
        Dim k(,)
        Dim m(5, 10)

        Dim o(,) As Integer = {{1, 2, 3}, {4, 5, 6}, {7, 8, 9}, {10, 11, 13}}
        Dim p(,) As Integer = {{1, 2, 3}, {4, 5, 6}, {7, 8, 9}, {10, 11}} ' Error, not rectangular
        Dim q(1, 2) As Integer = {{1, 2, 3}, {1, 2, 3}}                   ' Error, no explicit quantification

        ' or
        Dim r(,) As Integer = New Integer(,) {{1}, {2}, {3}}
        Dim s(,) As Integer = New Integer(2, 0) {{1}, {2}, {3}}
        Dim t = New Integer(1, 2) {{1, 2, 3}, {4, 5, 6}}

        ' explicit
        Dim u As Integer(,) = New Integer(1, 2) {{1, 2, 3}, {4, 5, 6}}
        Dim v(,) As Integer(,) = New Integer(1, 2) {{1, 2, 3}, {4, 5, 6}} ' only once
        Dim w(,) = New Integer(1, 0) {{1}, {2}}
        ReDim w(4, 4)

        ' sequential output
        For Each item In t
            WriteLine(item)
        Next

        ' rows and columns
        For i = 0 To t.GetLength(0) - 1
            For j = 0 To t.GetLength(1) - 1
                Write(t(i, j) & " ")
            Next
            WriteLine()
        Next

        ' Jagged array (nonrectangular)
        Dim x(10)(10) As Integer ' Error, only first dimension can be quantified
        Dim x()() As Integer
        Dim y As Integer()()
        Dim z%(10)()
        ReDim z(10)(10) ' both dimensions must be quantified and the first can't be > 10

        'Initialization
        x = New Integer(5)() {}
        x(0) = {1, 2, 3}
        x(1) = {1, 2}
        z(0) = {1}
        z(1) = {1, 2}
        z = {New Integer() {1, 2}, New Integer() {1, 2, 3}}

        For Each row As Integer() In x
            If IsNothing(row) Then Exit For
            For Each e As Integer In row
                Write(e & " ")
            Next
            WriteLine()
        Next

        ' or
        For i = 0 To x.GetLength(0) - 1
            If IsNothing(x(i)) Then Exit For
            For j = 0 To x(i).GetLength(0) - 1
                Write(x(i)(j) & " ")
            Next
            WriteLine()
        Next

        Dim tab(2, 4, 5) As UShort
        tab = {
                {
                  {8825, 7, 5, 3, 0, 0},
                  {1254, 0, 2, 2, 3, 0},
                  {4853, 4, 0, 5, 0, 0},
                  {2267, 0, 8, 8, 7, 0},
                  {0, 0, 0, 0, 0, 0}
                },
                {
                  {7458, 2, 3, 0, 0, 0},
                  {6462, 4, 2, 0, 0, 0},
                  {9652, 0, 0, 0, 1, 4},
                  {0, 0, 0, 0, 0, 0},
                  {0, 0, 0, 0, 0, 0}
                },
                {
                  {9132, 0, 0, 0, 3, 3},
                  {5977, 0, 0, 0, 2, 0},
                  {0, 0, 0, 0, 0, 0},
                  {0, 0, 0, 0, 0, 0},
                  {0, 0, 0, 0, 0, 0}
                }
              }
    End Sub
#End If

    Function GetNewLicenceKeys(Optional quantity = 10) As Guid()
        Dim keys(quantity) As Guid

        For index = 0 To quantity
            keys(index) = System.Guid.NewGuid()
        Next

        Return keys
    End Function
End Module
