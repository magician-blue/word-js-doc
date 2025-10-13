# Word.Interfaces.BuildingBlockLoadOptions interface

Package: https://learn.microsoft.com/en-us/javascript/api/word

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents a building block in a template. A building block is pre-built content, similar to autotext, that may contain text, images, and formatting.

## Remarks

[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- $all: Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
- category: Returns a `BuildingBlockCategory` object that represents the category for the building block.
- description: Specifies the description for the building block.
- id: Returns the internal identification number for the building block.
- index: Returns the position of this building block in a collection.
- insertType: Specifies a `DocPartInsertType` value that represents how to insert the contents of the building block into the document.
- name: Specifies the name of the building block.
- type: Returns a `BuildingBlockTypeItem` object that represents the type for the building block.
- value: Specifies the contents of the building block.

## Property Details

### $all

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

#### Property Value
boolean

---

### category

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `BuildingBlockCategory` object that represents the category for the building block.

```typescript
category?: Word.Interfaces.BuildingBlockCategoryLoadOptions;
```

#### Property Value
Word.Interfaces.BuildingBlockCategoryLoadOptions  
https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.buildingblockcategoryloadoptions

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### description

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the description for the building block.

```typescript
description?: boolean;
```

#### Property Value
boolean

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### id

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the internal identification number for the building block.

```typescript
id?: boolean;
```

#### Property Value
boolean

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### index

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the position of this building block in a collection.

```typescript
index?: boolean;
```

#### Property Value
boolean

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### insertType

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a `DocPartInsertType` value that represents how to insert the contents of the building block into the document.

```typescript
insertType?: boolean;
```

#### Property Value
boolean

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### name

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the name of the building block.

```typescript
name?: boolean;
```

#### Property Value
boolean

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### type

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `BuildingBlockTypeItem` object that represents the type for the building block.

```typescript
type?: Word.Interfaces.BuildingBlockTypeItemLoadOptions;
```

#### Property Value
Word.Interfaces.BuildingBlockTypeItemLoadOptions  
https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.buildingblocktypeitemloadoptions

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### value

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the contents of the building block.

```typescript
value?: boolean;
```

#### Property Value
boolean

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)