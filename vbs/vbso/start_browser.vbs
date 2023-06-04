set b = createobject( "internetexplorer.application" )

b.visible = true
b.left = 10
b.navigate "https://www.aa.com"
msgbox "quit"
b.quit
