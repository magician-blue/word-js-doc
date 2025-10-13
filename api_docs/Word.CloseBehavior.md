# Word.CloseBehavior enum

Package: [word](https://learn.microsoft.com/en-us/javascript/api/word)

Specifies the close behavior for `Document.close`.

## Remarks

[API set: WordApi 1.5](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/save-close.yaml

// Closes the document after saving.
await Word.run(async (context) => {
  context.document.close(Word.CloseBehavior.save);
});
```

## Fields

- save = "Save"
  - Saves the changes before closing the document.
  - [API set: WordApi 1.5](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- skipSave = "SkipSave"
  - Discard the possible changes when closing the document.
  - [API set: WordApi 1.5](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)