#import "@preview/touying:0.6.3": *
#import themes.simple: *

#show: simple-theme.with(aspect-ratio: "16-9")

= Code Block Tests

== Inline Code

Here is some `inline code` in a sentence.

Multiple inline: `foo`, `bar`, and `baz`.

== Code Block

```python
def hello():
    print("Hello, World!")

for i in range(10):
    hello()
```

== Another Language

```rust
fn main() {
    let x = 42;
    println!("The answer is {}", x);
}
```

== Mixed Content

Some text before the code block.

```javascript
const add = (a, b) => a + b;
console.log(add(1, 2));
```

And some text after the code block with `inline code` mixed in.
