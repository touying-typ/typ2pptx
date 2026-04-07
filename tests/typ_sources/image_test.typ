#import "@preview/touying:0.6.3": *
#import themes.simple: *

#show: simple-theme.with(aspect-ratio: "16-9")

= Image Embedding Tests

== PNG Image

Here is an embedded PNG image:

#image("test_image.png", width: 50%)

Text after the PNG image.

== SVG Image

Here is an embedded SVG image:

#image("test_vector.svg", width: 50%)

Text after the SVG image.

// Note: PDF images are NOT supported by typst-ts-cli (v0.6.0 gives
// "unknown image format"). PDF rasterization is tested separately
// via unit tests in TestImageRasterization.

== Multiple Images

#grid(
  columns: 2,
  gutter: 16pt,
  image("test_image.png", width: 100%),
  image("test_vector.svg", width: 100%),
)

Caption: Two images side by side.
