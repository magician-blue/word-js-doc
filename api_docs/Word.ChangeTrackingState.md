# Word.ChangeTrackingState enum

Package: [word](/en-us/javascript/api/word)

Specify the track state when ChangeTracking is on.

## Remarks

[ [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

#### Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/10-content-controls/get-change-tracking-states.yaml

// Logs the current change tracking states of the content controls.
await Word.run(async (context) => {
  let trackAddedArray: Word.ChangeTrackingState[] = [Word.ChangeTrackingState.added];
  let trackDeletedArray: Word.ChangeTrackingState[] = [Word.ChangeTrackingState.deleted];
  let trackNormalArray: Word.ChangeTrackingState[] = [Word.ChangeTrackingState.normal];

  let addedContentControls = context.document.body.getContentControls().getByChangeTrackingStates(trackAddedArray);
  let deletedContentControls = context.document.body
    .getContentControls()
    .getByChangeTrackingStates(trackDeletedArray);
  let normalContentControls = context.document.body.getContentControls().getByChangeTrackingStates(trackNormalArray);

  addedContentControls.load();
  deletedContentControls.load();
  normalContentControls.load();
  await context.sync();

  console.log(`Number of content controls in Added state: ${addedContentControls.items.length}`);
  console.log(`Number of content controls in Deleted state: ${deletedContentControls.items.length}`);
  console.log(`Number of content controls in Normal state: ${normalContentControls.items.length}`);
});
```

## Fields

- added = "Added" [ [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]
- deleted = "Deleted" [ [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]
- normal = "Normal" [ [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]
- unknown = "Unknown" [ [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]