# Word.OutlineLevel enum

Package: [word](https://learn.microsoft.com/en-us/javascript/api/word)

Represents the outline levels.

## Remarks

[API set: WordApi 1.5](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/40-tables/manage-custom-style.yaml

// Imports styles from JSON.
await Word.run(async (context) => {
  const str =
    '{"styles":[{"baseStyle":"Default Paragraph Font","builtIn":false,"inUse":true,"linked":false,"nameLocal":"NewCharStyle","priority":2,"quickStyle":true,"type":"Character","unhideWhenUsed":false,"visibility":false,"paragraphFormat":null,"font":{"name":"DengXian Light","size":16.0,"bold":true,"italic":false,"color":"#F1A983","underline":"None","subscript":false,"superscript":true,"strikeThrough":true,"doubleStrikeThrough":false,"highlightColor":null,"hidden":false},"shading":{"backgroundPatternColor":"#FF0000"}},{"baseStyle":"Normal","builtIn":false,"inUse":true,"linked":false,"nextParagraphStyle":"NewParaStyle","nameLocal":"NewParaStyle","priority":1,"quickStyle":true,"type":"Paragraph","unhideWhenUsed":false,"visibility":false,"paragraphFormat":{"alignment":"Centered","firstLineIndent":0.0,"keepTogether":false,"keepWithNext":false,"leftIndent":72.0,"lineSpacing":18.0,"lineUnitAfter":0.0,"lineUnitBefore":0.0,"mirrorIndents":false,"outlineLevel":"OutlineLevelBodyText","rightIndent":72.0,"spaceAfter":30.0,"spaceBefore":30.0,"widowControl":true},"font":{"name":"DengXian","size":14.0,"bold":true,"italic":true,"color":"#8DD873","underline":"Single","subscript":false,"superscript":false,"strikeThrough":false,"doubleStrikeThrough":true,"highlightColor":null,"hidden":false},"shading":{"backgroundPatternColor":"#00FF00"}},{"baseStyle":"Table Normal","builtIn":false,"inUse":true,"linked":false,"nextParagraphStyle":"NewTableStyle","nameLocal":"NewTableStyle","priority":100,"type":"Table","unhideWhenUsed":false,"visibility":false,"paragraphFormat":{"alignment":"Left","firstLineIndent":0.0,"keepTogether":false,"keepWithNext":false,"leftIndent":0.0,"lineSpacing":12.0,"lineUnitAfter":0.0,"lineUnitBefore":0.0,"mirrorIndents":false,"outlineLevel":"OutlineLevelBodyText","rightIndent":0.0,"spaceAfter":0.0,"spaceBefore":0.0,"widowControl":true},"font":{"name":"DengXian","size":20.0,"bold":false,"italic":true,"color":"#D86DCB","underline":"None","subscript":false,"superscript":false,"strikeThrough":false,"doubleStrikeThrough":false,"highlightColor":null,"hidden":false},"tableStyle":{"allowBreakAcrossPage":true,"alignment":"Left","bottomCellMargin":0.0,"leftCellMargin":0.08,"rightCellMargin":0.08,"topCellMargin":0.0,"cellSpacing":0.0},"shading":{"backgroundPatternColor":"#60CAF3"}}]}';
  const styles = context.document.importStylesFromJson(str);
  await context.sync();
  console.log("Styles imported from JSON:", styles);
});
```

## Fields

- outlineLevel1 = "OutlineLevel1"
  - Represents outline level 1.
  - [API set: WordApi 1.5](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- outlineLevel2 = "OutlineLevel2"
  - Represents outline level 2.
  - [API set: WordApi 1.5](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- outlineLevel3 = "OutlineLevel3"
  - Represents outline level 3.
  - [API set: WordApi 1.5](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- outlineLevel4 = "OutlineLevel4"
  - Represents outline level 4.
  - [API set: WordApi 1.5](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- outlineLevel5 = "OutlineLevel5"
  - Represents outline level 5.
  - [API set: WordApi 1.5](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- outlineLevel6 = "OutlineLevel6"
  - Represents outline level 6.
  - [API set: WordApi 1.5](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- outlineLevel7 = "OutlineLevel7"
  - Represents outline level 7.
  - [API set: WordApi 1.5](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- outlineLevel8 = "OutlineLevel8"
  - Represents outline level 8.
  - [API set: WordApi 1.5](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- outlineLevel9 = "OutlineLevel9"
  - Represents outline level 9.
  - [API set: WordApi 1.5](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- outlineLevelBodyText = "OutlineLevelBodyText"
  - Represents outline level body text, not an outline level.
  - [API set: WordApi 1.5](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)