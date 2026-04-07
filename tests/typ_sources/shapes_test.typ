#import "@preview/touying:0.6.3": *
#import themes.simple: *

#show: simple-theme.with(aspect-ratio: "16-9")

= Shapes Test

== Rectangles and Lines

#rect(width: 200pt, height: 80pt, fill: rgb("#4CAF50"), stroke: 2pt + black)[
  #align(center + horizon)[#text(fill: white, size: 18pt)[Green Box]]
]

#line(length: 300pt, stroke: 3pt + red)

#v(10pt)

#rect(width: 150pt, height: 60pt, fill: rgb("#2196F3"), radius: 10pt)

== Circles and Paths

#circle(radius: 40pt, fill: rgb("#FF9800"), stroke: 2pt + black)

#v(10pt)

#ellipse(width: 120pt, height: 60pt, fill: rgb("#9C27B0").lighten(30%))

#v(10pt)

#polygon(
  fill: rgb("#F44336").lighten(20%),
  stroke: 1pt + black,
  (0pt, 50pt),
  (50pt, 0pt),
  (100pt, 50pt),
  (75pt, 100pt),
  (25pt, 100pt),
)

== Images and Complex Shapes

#box(width: 200pt, height: 100pt, fill: gradient.linear(rgb("#1a237e"), rgb("#4fc3f7")))

#v(10pt)

#grid(columns: 3, gutter: 20pt,
  rect(width: 60pt, height: 60pt, fill: red),
  rect(width: 60pt, height: 60pt, fill: green),
  rect(width: 60pt, height: 60pt, fill: blue),
)

#v(10pt)

Text with shapes: #box(width: 20pt, height: 20pt, fill: orange, baseline: 5pt) inline.
