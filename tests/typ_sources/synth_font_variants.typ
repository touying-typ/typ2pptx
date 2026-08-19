// Network-independent font-variant fixture: no @preview package imports (unlike basic_text.typ, which needs touying from the Typst package registry). Uses only system fonts (Helvetica Neue proportional, Menlo monospace) so it compiles offline via typst-ts-cli or the equivalent node-compiler API. Covers regular/bold/italic/bolditalic/mono/monobold -- the six font_variant styles typst_svg_parser.py can produce.

#set page(width: 400pt, height: 300pt, margin: 20pt)
#set text(font: "Helvetica Neue", size: 14pt)

Regular text lorem ipsum dolor sit amet consectetur.

#text(weight: "bold")[Bold text lorem ipsum dolor sit amet consectetur.]

#text(style: "italic")[Italic text lorem ipsum dolor sit amet consectetur.]

#text(weight: "bold", style: "italic")[Bold italic text lorem ipsum dolor sit.]

#text(font: "Menlo")[Mono regular text lorem ipsum dolor sit amet.]

#text(font: "Menlo", weight: "bold")[Mono bold text lorem ipsum dolor sit amet.]
