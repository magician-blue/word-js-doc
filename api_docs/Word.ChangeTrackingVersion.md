# Word.ChangeTrackingVersion enum

Package: [word](/en-us/javascript/api/word)

Specify the current version or the original version of the text.

## Remarks

[API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-change-tracking.yaml

// Gets the reviewed text.
await Word.run(async (context) => {
  const range: Word.Range = context.document.getSelection();
  const before = range.getReviewedText(Word.ChangeTrackingVersion.original);
  const after = range.getReviewedText(Word.ChangeTrackingVersion.current);

  await context.sync();

  console.log("Reviewed text (before):", before.value, "Reviewed text (after):", after.value);
});
```

## Fields

- current = "Current"
  - [API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- original = "Original"
  - [API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)