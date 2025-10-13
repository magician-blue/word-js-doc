# Word.HeaderFooterType enum

Package: [word](/en-us/javascript/api/word)

## Remarks

[API set: WordApi 1.1]

#### Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/25-paragraph/insert-header-and-footer.yaml

await Word.run(async (context) => {
  context.document.sections
    .getFirst()
    .getHeader(Word.HeaderFooterType.primary)
    .insertParagraph("This is a primary header.", "End");

  await context.sync();
});
```

## Fields

- evenPages = "EvenPages"
  - Returns all headers or footers on even-numbered pages of a section.
  - [API set: WordApi 1.1]

- firstPage = "FirstPage"
  - Returns the header or footer on the first page of a section.
  - [API set: WordApi 1.1]

- primary = "Primary"
  - Returns the header or footer on all pages of a section, but excludes the first page or even pages if they are different.
  - [API set: WordApi 1.1]