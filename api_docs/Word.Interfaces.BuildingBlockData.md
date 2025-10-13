# Word.Interfaces.BuildingBlockData interface

An interface describing the data returned by calling `buildingBlock.toJSON()`.

- Package: [word](/en-us/javascript/api/word)

## Properties

- [description](#description)  
  Specifies the description for the building block.

- [id](#id)  
  Returns the internal identification number for the building block.

- [index](#index)  
  Returns the position of this building block in a collection.

- [insertType](#inserttype)  
  Specifies a `DocPartInsertType` value that represents how to insert the contents of the building block into the document.

- [name](#name)  
  Specifies the name of the building block.

- [value](#value)  
  Specifies the contents of the building block.

## Property Details

### description

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the description for the building block.

```typescript
description?: string;
```

Property value: string

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### id

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the internal identification number for the building block.

```typescript
id?: string;
```

Property value: string

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### index

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the position of this building block in a collection.

```typescript
index?: number;
```

Property value: number

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### insertType

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a `DocPartInsertType` value that represents how to insert the contents of the building block into the document.

```typescript
insertType?: Word.DocPartInsertType | "Content" | "Paragraph" | "Page";
```

Property value: [Word.DocPartInsertType](/en-us/javascript/api/word/word.docpartinserttype) | "Content" | "Paragraph" | "Page"

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### name

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the name of the building block.

```typescript
name?: string;
```

Property value: string

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### value

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the contents of the building block.

```typescript
value?: string;
```

Property value: string

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)