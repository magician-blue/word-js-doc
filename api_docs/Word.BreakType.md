# Word.BreakType enum

Package: [word](/en-us/javascript/api/word)

Specifies the form of a break.

## Remarks

[ [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

#### Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/25-paragraph/insert-line-and-page-breaks.yaml

await Word.run(async (context) => {
  context.document.body.paragraphs.getFirst().insertBreak(Word.BreakType.page, "After");

  await context.sync();
  console.log("success");
});
```

## Fields

- line = "Line"
  - Line break.
  - [ [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

- next = "Next"
  - Warning: next has been deprecated. Use sectionNext instead.
  - [ [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

- page = "Page"
  - Page break at the insertion point.
  - [ [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

- sectionContinuous = "SectionContinuous"
  - New section without a corresponding page break.
  - [ [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

- sectionEven = "SectionEven"
  - Section break with the next section beginning on the next even-numbered page. If the section break falls on an even-numbered page, Word leaves the next odd-numbered page blank.
  - [ [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

- sectionNext = "SectionNext"
  - Section break on next page.
  - [ [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

- sectionOdd = "SectionOdd"
  - Section break with the next section beginning on the next odd-numbered page. If the section break falls on an odd-numbered page, Word leaves the next even-numbered page blank.
  - [ [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]