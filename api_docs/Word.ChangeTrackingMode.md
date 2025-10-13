# Word.ChangeTrackingMode enum

Package: [word](/en-us/javascript/api/word)

Represents the possible change tracking modes.

## Remarks

[ API set: WordApi 1.4 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-change-tracking.yaml

// Gets the current change tracking mode.
await Word.run(async (context) => {
  const document: Word.Document = context.document;
  document.load("changeTrackingMode");
  await context.sync();

  if (document.changeTrackingMode === Word.ChangeTrackingMode.trackMineOnly) {
    console.log("Only my changes are being tracked.");
  } else if (document.changeTrackingMode === Word.ChangeTrackingMode.trackAll) {
    console.log("Everyone's changes are being tracked.");
  } else {
    console.log("No changes are being tracked.");
  }
});
```

## Fields

|  |  |
| --- | --- |
| off = "Off" | ChangeTracking is turned off.<br><br>- [ API set: WordApi 1.4 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) |
| trackAll = "TrackAll" | ChangeTracking is turned on for everyone.<br><br>- [ API set: WordApi 1.4 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) |
| trackMineOnly = "TrackMineOnly" | Tracking is turned on for my changes only.<br><br>- [ API set: WordApi 1.4 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) |