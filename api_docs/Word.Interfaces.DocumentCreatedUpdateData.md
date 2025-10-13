# Word.Interfaces.DocumentCreatedUpdateData interface

Package: [word](/en-us/javascript/api/word)

An interface for updating data on the DocumentCreated object, for use in `documentCreated.set({ ... })`.

## Properties

- body
  - Gets the body object of the document. The body is the text that excludes headers, footers, footnotes, textboxes, etc.
- properties
  - Gets the properties of the document.

## Property Details

### body

Gets the body object of the document. The body is the text that excludes headers, footers, footnotes, textboxes, etc.

```typescript
body?: Word.Interfaces.BodyUpdateData;
```

- Property value: [Word.Interfaces.BodyUpdateData](/en-us/javascript/api/word/word.interfaces.bodyupdatedata)

Remarks: [API set: WordApiHiddenDocument 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### properties

Gets the properties of the document.

```typescript
properties?: Word.Interfaces.DocumentPropertiesUpdateData;
```

- Property value: [Word.Interfaces.DocumentPropertiesUpdateData](/en-us/javascript/api/word/word.interfaces.documentpropertiesupdatedata)

Remarks: [API set: WordApiHiddenDocument 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)