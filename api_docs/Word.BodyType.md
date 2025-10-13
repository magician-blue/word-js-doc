# Word.BodyType enum

Package: https://learn.microsoft.com/en-us/javascript/api/word

Represents the types of body objects.

## Remarks

[API set: WordApi 1.3](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-footnotes.yaml

// Gets the referenced note's item type and body type, which are both "Footnote".
await Word.run(async (context) => {
  const footnotes: Word.NoteItemCollection = context.document.body.footnotes;
  footnotes.load("items");
  await context.sync();

  const referenceNumber = (document.getElementById("input-reference") as HTMLInputElement).value;
  const mark = (referenceNumber as number) - 1;
  const item: Word.NoteItem = footnotes.items[mark];
  console.log(`Note type of footnote ${referenceNumber}: ${item.type}`);

  item.body.load("type");
  await context.sync();

  console.log(`Body type of note: ${item.body.type}`);
});
```

## Fields

- endnote = "Endnote"
  - Endnote body.
  - [API set: WordApi 1.5](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- footer = "Footer"
  - Footer body.
  - [API set: WordApi 1.3](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- footnote = "Footnote"
  - Footnote body.
  - [API set: WordApi 1.5](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- header = "Header"
  - Header body.
  - [API set: WordApi 1.3](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- mainDoc = "MainDoc"
  - Main document body.
  - [API set: WordApi 1.3](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- noteItem = "NoteItem"
  - Note body e.g., endnote, footnote.
  - [API set: WordApi 1.5](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- section = "Section"
  - Section body.
  - [API set: WordApi 1.3](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- shape = "Shape"
  - Shape body.
  - [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- tableCell = "TableCell"
  - Table cell body.
  - [API set: WordApi 1.3](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- unknown = "Unknown"
  - Unknown body type.
  - [API set: WordApi 1.3](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)