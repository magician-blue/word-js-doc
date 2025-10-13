# Word.CritiquePopupOptions interface

Package: [word](/en-us/javascript/api/word)

Properties defining the behavior of the pop-up menu for a given critique.

## Remarks

[API set: WordApi 1.8](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### Examples

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

- brandingTextResourceId  
  Gets the manifest resource ID of the string to use for branding. This branding text appears next to your add-in icon in the pop-up menu.

- subtitleResourceId  
  Gets the manifest resource ID of the string to use as the subtitle.

- suggestions  
  Gets the suggestions to display in the critique pop-up menu.

- titleResourceId  
  Gets the manifest resource ID of the string to use as the title.

## Property Details

### brandingTextResourceId

Gets the manifest resource ID of the string to use for branding. This branding text appears next to your add-in icon in the pop-up menu.

```typescript
brandingTextResourceId: string;
```

Property Value: string

Remarks  
[API set: WordApi 1.8](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### subtitleResourceId

Gets the manifest resource ID of the string to use as the subtitle.

```typescript
subtitleResourceId: string;
```

Property Value: string

Remarks  
[API set: WordApi 1.8](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### suggestions

Gets the suggestions to display in the critique pop-up menu.

```typescript
suggestions: string[];
```

Property Value: string[]

Remarks  
[API set: WordApi 1.8](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### titleResourceId

Gets the manifest resource ID of the string to use as the title.

```typescript
titleResourceId: string;
```

Property Value: string

Remarks  
[API set: WordApi 1.8](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)