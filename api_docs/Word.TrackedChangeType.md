# Word.TrackedChangeType enum

- Package: [word](/en-us/javascript/api/word)

TrackedChange type.

## Remarks

[ [API set: WordApi 1.6](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

#### Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-tracked-changes.yaml

// Gets the next (second) tracked change.
await Word.run(async (context) => {
  const body: Word.Body = context.document.body;
  const trackedChanges: Word.TrackedChangeCollection = body.getTrackedChanges();
  await context.sync();

  const trackedChange: Word.TrackedChange = trackedChanges.getFirst();
  await context.sync();

  const nextTrackedChange: Word.TrackedChange = trackedChange.getNext();
  await context.sync();

  nextTrackedChange.load(["author", "date", "text", "type"]);
  await context.sync();

  console.log(nextTrackedChange);
});
```

## Fields

- added = "Added"
  - Add change.
  - [ [API set: WordApi 1.6](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

- deleted = "Deleted"
  - Delete change.
  - [ [API set: WordApi 1.6](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

- formatted = "Formatted"
  - Format change.
  - [ [API set: WordApi 1.6](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

- none = "None"
  - No revision.
  - [ [API set: WordApi 1.6](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]