# Word.InsertShapeOptions interface

Package: [word](/en-us/javascript/api/word)

Specifies the options to determine location and size when inserting a shape.

## Remarks

[ API set: WordApiDesktop 1.2 ]

#### Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-shapes-text-boxes.yaml

await Word.run(async (context) => {
  // Inserts a text box at the beginning of the selection.
  const range: Word.Range = context.document.getSelection();
  const insertShapeOptions: Word.InsertShapeOptions = {
    top: 0,
    left: 0,
    height: 100,
    width: 100
  };

  const newTextBox: Word.Shape = range.insertTextBox("placeholder text", insertShapeOptions);
  await context.sync();

  console.log("Inserted a text box at the beginning of the current selection.");
});
```

## Properties

- [height](#height)  
  Represents the height of the shape being inserted.
- [left](#left)  
  Represents the left position of the shape being inserted.
- [top](#top)  
  Represents the top position of the shape being inserted.
- [width](#width)  
  Represents the width of the shape being inserted.

## Property Details

### height

Represents the height of the shape being inserted.

```typescript
height?: number;
```

Property Value  
number

Remarks  
[ API set: WordApiDesktop 1.2 ]

### left

Represents the left position of the shape being inserted.

```typescript
left?: number;
```

Property Value  
number

Remarks  
[ API set: WordApiDesktop 1.2 ]

### top

Represents the top position of the shape being inserted.

```typescript
top?: number;
```

Property Value  
number

Remarks  
[ API set: WordApiDesktop 1.2 ]

### width

Represents the width of the shape being inserted.

```typescript
width?: number;
```

Property Value  
number

Remarks  
[ API set: WordApiDesktop 1.2 ]