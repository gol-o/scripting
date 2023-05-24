var f = function() { alert("first") }
f()

var g = function() { alert("second") }()
g() // no calling

var h = function() { alert("third") }
()
