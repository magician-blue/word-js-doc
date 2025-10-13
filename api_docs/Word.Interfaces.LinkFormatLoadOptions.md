# Word.Interfaces.LinkFormatLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents the linking characteristics for an OLE object or picture.

## Remarks

[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- $all  
  Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
- isAutoUpdated  
  Specifies if the link is updated automatically when the container file is opened or when the source file is changed.
- isLocked  
  Specifies if a `Field`, `InlineShape`, or `Shape` object is locked to prevent automatic updating.
- isPictureSavedWithDocument  
  Specifies if the linked picture is saved with the document.
- sourceFullName  
  Specifies the path and name of the source file for the linked OLE object, picture, or field.
- sourceName  
  Gets the name of the source file for the linked OLE object, picture, or field.
- sourcePath  
  Gets the path of the source file for the linked OLE object, picture, or field.
- type  
  Gets the link type.

## Property Details

### $all

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

Property Value: boolean

### isAutoUpdated

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the link is updated automatically when the container file is opened or when the source file is changed.

```typescript
isAutoUpdated?: boolean;
```

Property Value: boolean

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### isLocked

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if a `Field`, `InlineShape`, or `Shape` object is locked to prevent automatic updating.

```typescript
isLocked?: boolean;
```

Property Value: boolean

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### isPictureSavedWithDocument

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the linked picture is saved with the document.

```typescript
isPictureSavedWithDocument?: boolean;
```

Property Value: boolean

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### sourceFullName

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the path and name of the source file for the linked OLE object, picture, or field.

```typescript
sourceFullName?: boolean;
```

Property Value: boolean

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### sourceName

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the name of the source file for the linked OLE object, picture, or field.

```typescript
sourceName?: boolean;
```

Property Value: boolean

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### sourcePath

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the path of the source file for the linked OLE object, picture, or field.

```typescript
sourcePath?: boolean;
```

Property Value: boolean

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### type

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the link type.

```typescript
type?: boolean;
```

Property Value: boolean

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)