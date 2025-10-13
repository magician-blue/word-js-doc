# Word.Interfaces.LinkFormatData interface

Package: [word](/en-us/javascript/api/word)

An interface describing the data returned by calling `linkFormat.toJSON()`.

## Properties

- `isAutoUpdated` — Specifies if the link is updated automatically when the container file is opened or when the source file is changed.
- `isLocked` — Specifies if a `Field`, `InlineShape`, or `Shape` object is locked to prevent automatic updating.
- `isPictureSavedWithDocument` — Specifies if the linked picture is saved with the document.
- `sourceFullName` — Specifies the path and name of the source file for the linked OLE object, picture, or field.
- `sourceName` — Gets the name of the source file for the linked OLE object, picture, or field.
- `sourcePath` — Gets the path of the source file for the linked OLE object, picture, or field.
- `type` — Gets the link type.

## Property Details

### isAutoUpdated

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the link is updated automatically when the container file is opened or when the source file is changed.

```typescript
isAutoUpdated?: boolean;
```

Property Value  
boolean

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### isLocked

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if a `Field`, `InlineShape`, or `Shape` object is locked to prevent automatic updating.

```typescript
isLocked?: boolean;
```

Property Value  
boolean

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### isPictureSavedWithDocument

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the linked picture is saved with the document.

```typescript
isPictureSavedWithDocument?: boolean;
```

Property Value  
boolean

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### sourceFullName

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the path and name of the source file for the linked OLE object, picture, or field.

```typescript
sourceFullName?: string;
```

Property Value  
string

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### sourceName

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the name of the source file for the linked OLE object, picture, or field.

```typescript
sourceName?: string;
```

Property Value  
string

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### sourcePath

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the path of the source file for the linked OLE object, picture, or field.

```typescript
sourcePath?: string;
```

Property Value  
string

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### type

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the link type.

```typescript
type?: Word.LinkType | "Ole" | "Picture" | "Text" | "Reference" | "Include" | "Import" | "Dde" | "DdeAuto" | "Chart";
```

Property Value  
[`Word.LinkType`](/en-us/javascript/api/word/word.linktype) | "Ole" | "Picture" | "Text" | "Reference" | "Include" | "Import" | "Dde" | "DdeAuto" | "Chart"

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)