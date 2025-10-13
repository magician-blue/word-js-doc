# Word.NoteItemType enum

Package: [word](/en-us/javascript/api/word)

Note item type

## Remarks

[ [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

#### Examples

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
  - [ [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]
- footnote = "Footnote"
  - [ [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]