# Word.BorderType enum

Package: [word](/en-us/javascript/api/word)

## Remarks

[API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/40-tables/manage-formatting.yaml

// Gets border details about the first of the first table in the document.
await Word.run(async (context) => {
  const firstTable: Word.Table = context.document.body.tables.getFirst();
  const firstCell: Word.TableCell = firstTable.getCell(0, 0);
  const borderLocation = "Left";
  const border: Word.TableBorder = firstCell.getBorder(borderLocation);
  border.load(["type", "color", "width"]);
  await context.sync();

  console.log(
    `Details about the ${borderLocation} border of the first table's first cell:`,
    `- Color: ${border.color}`,
    `- Type: ${border.type}`,
    `- Width: ${border.width} points`
  );
});
```

## Fields

- dashDotStroked = "DashDotStroked"
  - [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- dashed = "Dashed"
  - [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- dashedSmall = "DashedSmall"
  - [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- dot2Dashed = "Dot2Dashed"
  - [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- dotDashed = "DotDashed"
  - [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- dotted = "Dotted"
  - [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- double = "Double"
  - [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- doubleWave = "DoubleWave"
  - [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- mixed = "Mixed"
  - [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- none = "None"
  - [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- single = "Single"
  - [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- thickThinLarge = "ThickThinLarge"
  - [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- thickThinMed = "ThickThinMed"
  - [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- thickThinSmall = "ThickThinSmall"
  - [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- thinThickLarge = "ThinThickLarge"
  - [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- thinThickMed = "ThinThickMed"
  - [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- thinThickSmall = "ThinThickSmall"
  - [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- thinThickThinLarge = "ThinThickThinLarge"
  - [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- thinThickThinMed = "ThinThickThinMed"
  - [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- thinThickThinSmall = "ThinThickThinSmall"
  - [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- threeDEmboss = "ThreeDEmboss"
  - [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- threeDEngrave = "ThreeDEngrave"
  - [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- triple = "Triple"
  - [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- wave = "Wave"
  - [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)