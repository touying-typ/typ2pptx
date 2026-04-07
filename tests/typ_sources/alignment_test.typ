#import "@preview/touying:0.6.3": *
#import themes.simple: *

#show: simple-theme.with(aspect-ratio: "16-9")

= Alignment Tests

== Left Aligned (Default)

This text is left aligned by default.

It should appear on the left side.

== Center Aligned

#align(center)[This text is centered on the slide.]

#align(center)[Another centered line.]

== Right Aligned

#align(right)[This text is right aligned.]

#align(right)[Another right aligned line.]

== Justified Text

#set par(justify: true)

#lorem(80)
