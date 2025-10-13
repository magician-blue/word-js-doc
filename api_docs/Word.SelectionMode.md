# Word.SelectionMode enum

Package: [word](/en-us/javascript/api/word)

This enum sets where the cursor (insertion point) in the document is after a selection.

## Remarks

[API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/35-ranges/scroll-to-range.yaml

await Word.run(async (context) => {
  // Select can be at the start or end of a range; this by definition moves the insertion point without selecting the range.
  context.document.body.paragraphs.getLast().select(Word.SelectionMode.end);

  await context.sync();
});
```

## Fields

- end = "End"
  - The cursor is at the end of the selection (just after the end of the selected range).
  - [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- select = "Select"
  - The entire range is selected.
  - [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- start = "Start"
  - The cursor is at the beginning of the selection (just before the start of the selected range).
  - [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)