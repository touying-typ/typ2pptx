#import "@preview/pinit:0.2.2": *
#import "@preview/touying:0.6.3": *
#import themes.simple: *

#show: simple-theme.with(aspect-ratio: "16-9")

= Algorithm Complexity

== Asymptotic Notation

#set text(size: 20pt)

For an algorithm with input size $n$:

- #pin(1)Constant: $O(1)$#pin(2) - Lookup in hash table
- #pin(3)Logarithmic: $O(log n)$#pin(4) - Binary search
- #pin(5)Linear: $O(n)$#pin(6) - Simple search
- #pin(7)Quadratic: $O(n^2)$#pin(8) - Bubble sort

#pinit-highlight(1, 2, fill: rgb(0, 180, 0, 40))
#pinit-highlight(3, 4, fill: rgb(0, 0, 180, 40))
#pinit-highlight(5, 6, fill: rgb(180, 180, 0, 40))
#pinit-highlight(7, 8, fill: rgb(180, 0, 0, 40))

== Pin Highlights

#set text(size: 20pt)

We can highlight important parts:

$f(x) = #pin(11)x^2#pin(12) + #pin(13)2x#pin(14) + 1$

#pinit-highlight(11, 12, fill: rgb(255, 0, 0, 40))
#pinit-highlight(13, 14, fill: rgb(0, 0, 255, 40))

The term #pin(15)$x^2$#pin(16) dominates for large $x$.

#pinit-highlight(15, 16, fill: rgb(255, 0, 0, 40))
