# Word.InsertLocation enum

Package: [word](/en-us/javascript/api/word)

The insertion location types.

## Remarks

[API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

To be used with an API call, such as `obj.insertSomething(newStuff, location);`. If the location is "Before" or "After", the new content will be outside of the modified object. If the location is "Start" or "End", the new content will be included as part of the modified object.

### Examples

```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/insert-section-breaks.yaml

// Inserts a section without an associated page break.
await Word.run(async (context) => {
  const body: Word.Body = context.document.body;
  body.insertBreak(Word.BreakType.sectionContinuous, Word.InsertLocation.end);

  await context.sync();

  console.log("Inserted section without an associated page break.");
});
```

## Fields

- after = "After"
  - Add content after the contents of the calling object.
  - [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- before = "Before"
  - Add content before the contents of the calling object.
  - [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- end = "End"
  - Append content to the contents of the calling object.
  - [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- replace = "Replace"
  - Replace the contents of the current object.
  - [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- start = "Start"
  - Prepend content to the contents of the calling object.
  - [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)