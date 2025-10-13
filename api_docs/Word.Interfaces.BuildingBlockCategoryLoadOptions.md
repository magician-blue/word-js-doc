# Word.Interfaces.BuildingBlockCategoryLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents a category of building blocks in a Word document.

## Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties
- $all: Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
- index: Returns the position of the `BuildingBlockCategory` object in a collection.
- name: Returns the name of the `BuildingBlockCategory` object.
- type: Returns a `BuildingBlockTypeItem` object that represents the type of building block for the building block category.

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

### index
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the position of the `BuildingBlockCategory` object in a collection.

```typescript
index?: boolean;
```

#### Property Value
boolean

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### name
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the name of the `BuildingBlockCategory` object.

```typescript
name?: boolean;
```

#### Property Value
boolean

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### type
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `BuildingBlockTypeItem` object that represents the type of building block for the building block category.

```typescript
type?: Word.Interfaces.BuildingBlockTypeItemLoadOptions;
```

#### Property Value
[Word.Interfaces.BuildingBlockTypeItemLoadOptions](/en-us/javascript/api/word/word.interfaces.buildingblocktypeitemloadoptions)

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)