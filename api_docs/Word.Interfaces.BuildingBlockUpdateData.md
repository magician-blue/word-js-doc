# Word.Interfaces.BuildingBlockUpdateData interface

Package: [word](https://learn.microsoft.com/en-us/javascript/api/word)

An interface for updating data on the BuildingBlock object, for use in `buildingBlock.set({ ... })`.

## Properties

- description  
  Specifies the description for the building block.
- insertType  
  Specifies a `DocPartInsertType` value that represents how to insert the contents of the building block into the document.
- name  
  Specifies the name of the building block.
- value  
  Specifies the contents of the building block.

## Property Details

### description

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the description for the building block.

```typescript
description?: string;
```

Property value: string

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### insertType

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a `DocPartInsertType` value that represents how to insert the contents of the building block into the document.

```typescript
insertType?: Word.DocPartInsertType | "Content" | "Paragraph" | "Page";
```

Property value: [Word.DocPartInsertType](https://learn.microsoft.com/en-us/javascript/api/word/word.docpartinserttype) | "Content" | "Paragraph" | "Page"

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### name

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the name of the building block.

```typescript
name?: string;
```

Property value: string

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### value

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the contents of the building block.

```typescript
value?: string;
```

Property value: string

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)