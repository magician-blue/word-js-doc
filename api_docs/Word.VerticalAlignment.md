# Word.VerticalAlignment enum

Package: [word](/en-us/javascript/api/word)

## Remarks
[API set: WordApi 1.3]

### Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/40-tables/manage-formatting.yaml

// Gets content alignment details about the first row of the first table in the document.
await Word.run(async (context) => {
  const firstTable: Word.Table = context.document.body.tables.getFirst();
  const firstTableRow: Word.TableRow = firstTable.rows.getFirst();
  firstTableRow.load(["horizontalAlignment", "verticalAlignment"]);
  await context.sync();

  console.log(`Details about the alignment of the first table's first row:`, `- Horizontal alignment of every cell in the row: ${firstTableRow.horizontalAlignment}`, `- Vertical alignment of every cell in the row: ${firstTableRow.verticalAlignment}`);
});
```

## Fields
- bottom = "Bottom"
  - [API set: WordApi 1.3]
- center = "Center"
  - [API set: WordApi 1.3]
- mixed = "Mixed"
  - [API set: WordApi 1.3]
- top = "Top"
  - [API set: WordApi 1.3]