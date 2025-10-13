# Word.ShapeType enum

- Package: [word](/en-us/javascript/api/word)

Represents the shape type.

## Remarks

[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples

```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-shapes-text-boxes.yaml

await Word.run(async (context) => {
  // Gets text boxes in main document.
  const shapes: Word.ShapeCollection = context.document.body.shapes;
  shapes.load();
  await context.sync();

  if (shapes.items.length > 0) {
    shapes.items.forEach(function(shape, index) {
      if (shape.type === Word.ShapeType.textBox) {
        console.log(`Shape ${index} in the main document has a text box. Properties:`, shape);
      }
    });
  } else {
    console.log("No shapes found in main document.");
  }
});
```

## Fields

- canvas = "Canvas"
  - Canvas shape.
  - [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- geometricShape = "GeometricShape"
  - Geometric shape.
  - [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- group = "Group"
  - Group shape.
  - [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- picture = "Picture"
  - Picture shape.
  - [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- textBox = "TextBox"
  - Text box shape.
  - [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- unsupported = "Unsupported"
  - Unsupported shape type.
  - [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)