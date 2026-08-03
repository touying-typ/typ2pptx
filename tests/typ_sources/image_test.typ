#import "@preview/touying:0.7.3": *
#import themes.simple: *

#show: simple-theme.with(aspect-ratio: "16-9")

#let test-png = bytes((
  137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82,
  0, 0, 0, 2, 0, 0, 0, 2, 8, 2, 0, 0, 0, 253, 212, 154, 115,
  0, 0, 0, 20, 73, 68, 65, 84, 120, 156, 99, 248, 207, 192, 192,
  0, 194, 12, 255, 255, 255, 103, 0, 0, 30, 239, 4, 252, 163, 200,
  180, 247, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130,
))
#let test-vector = bytes(
  "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"200\" height=\"120\" viewBox=\"0 0 200 120\"><rect width=\"200\" height=\"120\" rx=\"16\" fill=\"#2563eb\"/><circle cx=\"62\" cy=\"60\" r=\"34\" fill=\"#facc15\"/><path d=\"M112 34h58v14h-58zm0 28h42v14h-42z\" fill=\"white\"/></svg>"
)

= Image Embedding Tests

== PNG Image

Here is an embedded PNG image:

#image(test-png, format: "png", width: 50%)

Text after the PNG image.

== SVG Image

Here is an embedded SVG image:

#image(test-vector, format: "svg", width: 50%)

Text after the SVG image.

// Note: PDF images are NOT supported by typst-ts-cli (v0.6.0 gives
// "unknown image format"). PDF rasterization is tested separately
// via unit tests in TestImageRasterization.

== Multiple Images

#grid(
  columns: 2,
  gutter: 16pt,
  image(test-png, format: "png", width: 100%),
  image(test-vector, format: "svg", width: 100%),
)

Caption: Two images side by side.
