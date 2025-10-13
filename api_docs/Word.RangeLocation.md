# Word.RangeLocation enum

Package: [word](https://learn.microsoft.com/en-us/javascript/api/word)

Represents the location of a range. You can get range by calling getRange on different objects such as [Word.Paragraph](https://learn.microsoft.com/en-us/javascript/api/word/word.paragraph) and [Word.ContentControl](https://learn.microsoft.com/en-us/javascript/api/word/word.contentcontrol).

## Remarks

[ [API set: WordApi 1.3](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

## Examples

```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/25-paragraph/get-paragraph-on-insertion-point.yaml

await Word.run(async (context) => {
  // Get the complete sentence (as range) associated with the insertion point.
  const sentences: Word.RangeCollection = context.document
    .getSelection()
    .getTextRanges(["."] /* Using the "." as delimiter */, false /*means without trimming spaces*/);
  sentences.load("$none");
  await context.sync();

  // Expand the range to the end of the paragraph to get all the complete sentences.
  const sentencesToTheEndOfParagraph: Word.RangeCollection = sentences.items[0]
    .getRange()
    .expandTo(
      context.document
        .getSelection()
        .paragraphs.getFirst()
        .getRange(Word.RangeLocation.end)
    )
    .getTextRanges(["."], false /* Don't trim spaces*/);
  sentencesToTheEndOfParagraph.load("text");
  await context.sync();

  for (let i = 0; i < sentencesToTheEndOfParagraph.items.length; i++) {
    console.log(sentencesToTheEndOfParagraph.items[i].text);
  }
});
```

## Fields

- after = "After"
  - The point after the object. If the object is a paragraph content control or table content control, it's the point after the EOP or Table characters.
  - [API set: WordApi 1.3](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- before = "Before"
  - For content control only. It's the point before the opening tag.
  - [API set: WordApi 1.3](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- content = "Content"
  - The range between 'Start' and 'End'.
  - [API set: WordApi 1.3](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- end = "End"
  - The ending point of the object. For paragraph, it's the point before the EOP (end of paragraph). For content control, it's the point before the closing tag.
  - [API set: WordApi 1.3](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- start = "Start"
  - The starting point of the object. For content control, it's the point after the opening tag.
  - [API set: WordApi 1.3](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- whole = "Whole"
  - The object's whole range. If the object is a paragraph content control or table content control, the EOP or Table characters after the content control are also included.
  - [API set: WordApi 1.3](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)