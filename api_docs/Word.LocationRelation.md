# Word.LocationRelation enum

Package: [word](/en-us/javascript/api/word)

## Remarks

[API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### Examples

```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/35-ranges/compare-location.yaml

// Compares the location of one paragraph in relation to another paragraph.
await Word.run(async (context) => {
  const paragraphs: Word.ParagraphCollection = context.document.body.paragraphs;
  paragraphs.load("items");

  await context.sync();

  const firstParagraphAsRange: Word.Range = paragraphs.items[0].getRange();
  const secondParagraphAsRange: Word.Range = paragraphs.items[1].getRange();

  const comparedLocation = firstParagraphAsRange.compareLocationWith(secondParagraphAsRange);

  await context.sync();

  const locationValue: Word.LocationRelation = comparedLocation.value;
  console.log(`Location of the first paragraph in relation to the second paragraph: ${locationValue}`);
});
```

## Fields

- adjacentAfter = "AdjacentAfter"
  - Indicates that this instance occurs after, and is adjacent to, the range.
  - [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- adjacentBefore = "AdjacentBefore"
  - Indicates that this instance occurs before, and is adjacent to, the range.
  - [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- after = "After"
  - Indicates that this instance occurs after the range.
  - [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- before = "Before"
  - Indicates that this instance occurs before the range.
  - [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- contains = "Contains"
  - Indicates that this instance contains the range, with the exception of the start and end character of this instance.
  - [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- containsEnd = "ContainsEnd"
  - Indicates that this instance contains the range and that it shares the same end character. The range doesn't share the same start character as this instance.
  - [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- containsStart = "ContainsStart"
  - Indicates that this instance contains the range and that it shares the same start character. The range doesn't share the same end character as this instance.
  - [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- equal = "Equal"
  - Indicates that this instance and the range represent the same range.
  - [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- inside = "Inside"
  - Indicates that this instance is inside the range. The range doesn't share the same start and end characters as this instance.
  - [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- insideEnd = "InsideEnd"
  - Indicates that this instance is inside the range and that it shares the same end character. The range doesn't share the same start character as this instance.
  - [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- insideStart = "InsideStart"
  - Indicates that this instance is inside the range and that it shares the same start character. The range doesn't share the same end character as this instance.
  - [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- overlapsAfter = "OverlapsAfter"
  - Indicates that this instance starts inside the range and overlaps the range’s last character.
  - [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- overlapsBefore = "OverlapsBefore"
  - Indicates that this instance starts before the range and overlaps the range's first character.
  - [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- unrelated = "Unrelated"
  - Indicates that this instance and the range are in different sub-documents.
  - [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)