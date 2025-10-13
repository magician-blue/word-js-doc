# Word.InsertFileOptions interface

Package: [word](/en-us/javascript/api/word)

Specifies the options to determine what to copy when inserting a file.

## Remarks
[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### Examples
```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/insert-external-document.yaml

// Inserts content (applying selected settings) from another document passed in as a Base64-encoded string.
await Word.run(async (context) => {
  // Use the Base64-encoded string representation of the selected .docx file.
  context.document.insertFileFromBase64(externalDocument, "Replace", {
    importTheme: true,
    importStyles: true,
    importParagraphSpacing: true,
    importPageColor: true,
    importChangeTrackingMode: true,
    importCustomProperties: true,
    importCustomXmlParts: true,
    importDifferentOddEvenPages: true
  });
  await context.sync();
});
```

## Properties
- `importChangeTrackingMode`: Represents whether the change tracking mode status from the source document should be imported.
- `importCustomProperties`: Represents whether the custom properties from the source document should be imported. Overwrites existing properties with the same name.
- `importCustomXmlParts`: Represents whether the custom XML parts from the source document should be imported.
- `importDifferentOddEvenPages`: Represents whether to import the Different Odd and Even Pages setting for the header and footer from the source document.
- `importPageColor`: Represents whether the page color and other background information from the source document should be imported.
- `importParagraphSpacing`: Represents whether the paragraph spacing from the source document should be imported.
- `importStyles`: Represents whether the styles from the source document should be imported.
- `importTheme`: Represents whether the theme from the source document should be imported.

## Property Details

### importChangeTrackingMode
Represents whether the change tracking mode status from the source document should be imported.

```typescript
importChangeTrackingMode?: boolean;
```

Property Value: boolean

Remarks
[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### importCustomProperties
Represents whether the custom properties from the source document should be imported. Overwrites existing properties with the same name.

```typescript
importCustomProperties?: boolean;
```

Property Value: boolean

Remarks
[API set: WordApi 1.6](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### importCustomXmlParts
Represents whether the custom XML parts from the source document should be imported.

```typescript
importCustomXmlParts?: boolean;
```

Property Value: boolean

Remarks
[API set: WordApi 1.6](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### importDifferentOddEvenPages
Represents whether to import the Different Odd and Even Pages setting for the header and footer from the source document.

```typescript
importDifferentOddEvenPages?: boolean;
```

Property Value: boolean

Remarks
[API set: WordApi 1.7](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### importPageColor
Represents whether the page color and other background information from the source document should be imported.

```typescript
importPageColor?: boolean;
```

Property Value: boolean

Remarks
[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### importParagraphSpacing
Represents whether the paragraph spacing from the source document should be imported.

```typescript
importParagraphSpacing?: boolean;
```

Property Value: boolean

Remarks
[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### importStyles
Represents whether the styles from the source document should be imported.

```typescript
importStyles?: boolean;
```

Property Value: boolean

Remarks
[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### importTheme
Represents whether the theme from the source document should be imported.

```typescript
importTheme?: boolean;
```

Property Value: boolean

Remarks
[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)