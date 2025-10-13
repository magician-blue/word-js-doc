# Word.Alignment enum

Package: [word](https://learn.microsoft.com/en-us/javascript/api/word)

## Remarks

[ [API set: WordApi 1.1](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

#### Examples
```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/25-paragraph/paragraph-properties.yaml

await Word.run(async (context) => {
  const body: Word.Body = context.document.body;
  body.clear();
  body.insertParagraph(
    "Do you want to create a solution that extends the functionality of Word? You can use the Office Add-ins platform to extend Word clients running on the web, on a Windows desktop, or on a Mac.",
    "Start"
  );
  body.paragraphs
    .getLast()
    .insertText(
      "Use add-in commands to extend the Word UI and launch task panes that run JavaScript that interacts with the content in a Word document. Any code that you can run in a browser can run in a Word add-in. Add-ins that interact with content in a Word document create requests to act on Word objects and synchronize object state.",
      "Replace"
    );
  body.paragraphs.getFirst().alignment = "Left";
  body.paragraphs.getLast().alignment = Word.Alignment.left;
});
```

## Fields

- centered = "Centered"
  - Alignment to the center.
  - [ [API set: WordApi 1.1](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

- justified = "Justified"
  - Fully justified alignment.
  - [ [API set: WordApi 1.1](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

- left = "Left"
  - Alignment to the left.
  - [ [API set: WordApi 1.1](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

- mixed = "Mixed"
  - [ [API set: WordApi 1.1](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

- right = "Right"
  - Alignment to the right.
  - [ [API set: WordApi 1.1](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

- unknown = "Unknown"
  - Unknown alignment.
  - [ [API set: WordApi 1.1](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]