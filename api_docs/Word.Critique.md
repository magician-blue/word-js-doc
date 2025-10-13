# Word.Critique interface

Package: [word](/en-us/javascript/api/word)

Critique that will be rendered as underline for the specified part of paragraph in the document.

## Remarks

[API set: WordApi 1.7](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-annotations.yaml

// Adds annotations to the selected paragraph.
await Word.run(async (context) => {
  const paragraph: Word.Paragraph = context.document.getSelection().paragraphs.getFirst();
  const options: Word.CritiquePopupOptions = {
    brandingTextResourceId: "PG.TabLabel",
    subtitleResourceId: "PG.HelpCommand.TipTitle",
    titleResourceId: "PG.HelpCommand.Label",
    suggestions: ["suggestion 1", "suggestion 2", "suggestion 3"]
  };
  const critique1: Word.Critique = {
    colorScheme: Word.CritiqueColorScheme.red,
    start: 1,
    length: 3,
    popupOptions: options
  };
  const critique2: Word.Critique = {
    colorScheme: Word.CritiqueColorScheme.green,
    start: 6,
    length: 1,
    popupOptions: options
  };
  const critique3: Word.Critique = {
    colorScheme: Word.CritiqueColorScheme.blue,
    start: 10,
    length: 3,
    popupOptions: options
  };
  const critique4: Word.Critique = {
    colorScheme: Word.CritiqueColorScheme.lavender,
    start: 14,
    length: 3,
    popupOptions: options
  };
  const critique5: Word.Critique = {
    colorScheme: Word.CritiqueColorScheme.berry,
    start: 18,
    length: 10,
    popupOptions: options
  };
  const annotationSet: Word.AnnotationSet = {
    critiques: [critique1, critique2, critique3, critique4, critique5]
  };

  const annotationIds = paragraph.insertAnnotations(annotationSet);

  await context.sync();

  console.log("Annotations inserted:", annotationIds.value);
});
```

## Properties

- colorScheme: Specifies the color scheme of the critique.
- length: Specifies the length of the critique inside paragraph.
- popupOptions: Specifies the behavior of the pop-up menu for the critique.
- start: Specifies the start index of the critique inside paragraph.

## Property Details

### colorScheme

Specifies the color scheme of the critique.

```typescript
colorScheme: Word.CritiqueColorScheme | "Red" | "Green" | "Blue" | "Lavender" | "Berry";
```

Property Value: [Word.CritiqueColorScheme](/en-us/javascript/api/word/word.critiquecolorscheme) | "Red" | "Green" | "Blue" | "Lavender" | "Berry"

Remarks: [API set: WordApi 1.7](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### length

Specifies the length of the critique inside paragraph.

```typescript
length: number;
```

Property Value: number

Remarks: [API set: WordApi 1.7](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### popupOptions

Specifies the behavior of the pop-up menu for the critique.

```typescript
popupOptions?: Word.CritiquePopupOptions;
```

Property Value: [Word.CritiquePopupOptions](/en-us/javascript/api/word/word.critiquepopupoptions)

Remarks: [API set: WordApi 1.8](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### start

Specifies the start index of the critique inside paragraph.

```typescript
start: number;
```

Property Value: number

Remarks: [API set: WordApi 1.7](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)