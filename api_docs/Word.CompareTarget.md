# Word.CompareTarget enum

Package: [word](/en-us/javascript/api/word)

Specifies the target document for displaying document comparison differences.

## Remarks
[ [API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

#### Examples
```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/compare-documents.yaml

// Compares the current document with a specified external document.
await Word.run(async (context) => {
  // Absolute path of an online or local document.
  const filePath = (document.getElementById("filePath") as HTMLInputElement).value;
  // Options that configure the compare operation.
  const options: Word.DocumentCompareOptions = {
    compareTarget: Word.CompareTarget.compareTargetCurrent,
    detectFormatChanges: false
    // Other options you choose...
    };
  context.document.compare(filePath, options);

  await context.sync();

  console.log("Differences shown in the current document.");
});
```

## Fields

- compareTargetCurrent = "CompareTargetCurrent"
  - Places comparison differences in the current document.
  - [ [API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

- compareTargetNew = "CompareTargetNew"
  - Places comparison differences in a new document.
  - [ [API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

- compareTargetSelected = "CompareTargetSelected"
  - Places comparison differences in the target document.
  - [ [API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]