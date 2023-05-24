// calling anonymous function by function reference
var f = function() { console.log("f") }
f() // → f

// "iife" function (immediately invoked function expression)
// g assigned to the return value of the function
// alert() doesn't return a value
var g = function() { console.log("g") }()
g() // → error, g is not a function
console.log(g) // → undefined

// omitting ; sybmol may cause some irregular evaluation order
var h = function() { console.log("h") }
(function() { console.log("iife") })()
// iife function is interpreted as an (unused) argument of the previous function
