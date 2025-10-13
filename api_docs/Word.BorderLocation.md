# Word.BorderLocation enum

Package: [word](https://learn.microsoft.com/en-us/javascript/api/word)

## Remarks

[API set: WordApi 1.3](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples

```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/40-tables/manage-formatting.yaml

// Gets border details about the first table in the document.
await Word.run(async (context) => {
  const firstTable: Word.Table = context.document.body.tables.getFirst();
  const borderLocation = Word.BorderLocation.top;
  const border: Word.TableBorder = firstTable.getBorder(borderLocation);
  border.load(["type", "color", "width"]);
  await context.sync();

  console.log(`Details about the ${borderLocation} border of the first table:`, `- Color: ${border.color}`, `- Type: ${border.type}`, `- Width: ${border.width} points`);
});
```

## Fields

- all = "All"
  - [API set: WordApi 1.3](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- bottom = "Bottom"
  - [API set: WordApi 1.3](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- inside = "Inside"
  - [API set: WordApi 1.3](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- insideHorizontal = "InsideHorizontal"
  - [API set: WordApi 1.3](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- insideVertical = "InsideVertical"
  - [API set: WordApi 1.3](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- left = "Left"
  - [API set: WordApi 1.3](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- outside = "Outside"
  - [API set: WordApi 1.3](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- right = "Right"
  - [API set: WordApi 1.3](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- top = "Top"
  - [API set: WordApi 1.3](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)