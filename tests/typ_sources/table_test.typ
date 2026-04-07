#import "@preview/touying:0.6.3": *
#import themes.simple: *

#show: simple-theme.with(aspect-ratio: "16-9")

= Table Tests

== Simple Table

#table(
  columns: 3,
  [Name], [Age], [City],
  [Alice], [30], [New York],
  [Bob], [25], [London],
  [Charlie], [35], [Tokyo],
)

== Styled Table

#table(
  columns: (1fr, 2fr, 1fr),
  align: center,
  table.header(
    [*ID*], [*Description*], [*Status*],
  ),
  [1], [First item with longer text], [Active],
  [2], [Second item], [Inactive],
  [3], [Third item with even longer description text], [Pending],
)

== Table with Colors

#table(
  columns: 4,
  fill: (x, y) => if y == 0 { rgb("#4472C4") } else if calc.rem(y, 2) == 0 { rgb("#D6E4F0") },
  table.header(
    text(fill: white)[*Quarter*],
    text(fill: white)[*Revenue*],
    text(fill: white)[*Expenses*],
    text(fill: white)[*Profit*],
  ),
  [Q 1], [100k], [80k], [20k],
  [Q 2], [120k], [85k], [35k],
  [Q 3], [110k], [90k], [20k],
  [Q 4], [150k], [95k], [55k],
)

== Mixed Content Table

Text before table.

#table(
  columns: 2,
  [*Feature*], [*Supported*],
  [Bold text], [Yes],
  [Math: $x^2$], [Yes],
  [Colors], [#text(fill: red)[Red] / #text(fill: blue)[Blue]],
)

Text after table with inline math: $a + b = c$.
