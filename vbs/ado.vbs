dim rs 'the recordset used to sort keys  must be global
Set D = CreateObject("Scripting.Dictionary") 
for i=1 to 10
d.add right("0000"&Cint(rnd*10000),4), i
next

'
for each j in d.keys
   wscript.echo j & " " & d(j) 
next
wscript.echo ""

i=0
do
  b= DicNextItem(d,i)
  wscript.echo b(0)&" "&b(1)
loop until i=-1 

'---------------------------------------------

Function DicNextItem(dic,i) 
'returns successive items from dictionary ordered by keys
'arguments  dic: the dictionary
'         i: 0 must be passed at fist call, 
'                 returns 1 if there are more items 
'                 returns-1 if no more items   
'returns array with the key in index 0 and the value and value in index 1
'requires rs as global variable (no static in vbs)  
'it supposes the key is a string
      const advarchar=200
      const adopenstatic=3
      dim a(2)
      if i=0 then
        Set rs = CreateObject("ADODB.RECORDSET")
        with rs 
        .fields.append "Key", adVarChar, 100
        .CursorType = adOpenStatic
        .open
        'add all keys to the disconnected recordset  
        for each i in Dic.keys
          .AddNew
          rs("Key").Value = i
          .Update
        next
        .Sort= " Key ASC"
        .MoveFirst
        end with
        i=1
       end if
       if rs.EOF then 
         a=array(nul,nul)
       else
         a(0)=rs(0)
     a(1)=dic(a(0))
         rs.movenext
       end if
       if rs.EOF then i=-1:set rs=nothing
       DicNextItem=a
end function
