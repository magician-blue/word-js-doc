# Word.CellPaddingLocation enum

Package: [word](https://learn.microsoft.com/en-us/javascript/api/word)

## Remarks

[API set: WordApi 1.3]

### Examples

```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/40-tables/manage-formatting.yaml

// Gets cell padding details about the first table in the document.
await Word.run(async (context) => {
  const firstTable: Word.Table = context.document.body.tables.getFirst();
  const cellPaddingLocation = Word.CellPaddingLocation.right;
  const cellPadding = firstTable.getCellPadding(cellPaddingLocation);
  await context.sync();

  console.log(
    `Cell padding details about the ${cellPaddingLocation} border of the first table: ${cellPadding.value} points`
  );
});
```

## Fields

- bottom = "Bottom"
  - [API set: WordApi 1.3](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- left = "Left"
  - [API set: WordApi 1.3](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- right = "Right"
  - [API set: WordApi 1.3](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- top = "Top"
  - [API set: WordApi 1.3](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)