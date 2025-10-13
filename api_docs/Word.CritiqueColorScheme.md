# Word.CritiqueColorScheme enum

Package: [word](/en-us/javascript/api/word)

Represents the color scheme of a critique in the document, affecting underline and highlight.

## Remarks

[ [API set: WordApi 1.7](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

#### Examples
```TypeScript
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

## Fields

- berry = "Berry"
  - Berry color.
  - [ [API set: WordApi 1.7](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]
- blue = "Blue"
  - Blue color.
  - [ [API set: WordApi 1.7](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]
- green = "Green"
  - Green color.
  - [ [API set: WordApi 1.7](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]
- lavender = "Lavender"
  - Lavender color.
  - [ [API set: WordApi 1.7](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]
- red = "Red"
  - Red color.
  - [ [API set: WordApi 1.7](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]