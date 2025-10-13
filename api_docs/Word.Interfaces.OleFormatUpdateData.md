# Word.Interfaces.OleFormatUpdateData interface

- Package: [word](https://learn.microsoft.com/en-us/javascript/api/word)

An interface for updating data on the OleFormat object, for use in oleFormat.set({ ... }).

## Properties

- classType — Specifies the class type for the specified OLE object, picture, or field.
- iconIndex — Specifies the icon that is used when the displayAsIcon property is true.
- iconLabel — Specifies the text displayed below the icon for the OLE object.
- iconName — Specifies the program file in which the icon for the OLE object is stored.
- isFormattingPreservedOnUpdate — Specifies whether formatting done in Microsoft Word to the linked OLE object is preserved.

## Property Details

### classType

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the class type for the specified OLE object, picture, or field.

```typescript
classType?: string;
```

Property Value  
string

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### iconIndex

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the icon that is used when the displayAsIcon property is true.

```typescript
iconIndex?: number;
```

Property Value  
number

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### iconLabel

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the text displayed below the icon for the OLE object.

```typescript
iconLabel?: string;
```

Property Value  
string

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### iconName

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the program file in which the icon for the OLE object is stored.

```typescript
iconName?: string;
```

Property Value  
string

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### isFormattingPreservedOnUpdate

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether formatting done in Microsoft Word to the linked OLE object is preserved.

```typescript
isFormattingPreservedOnUpdate?: boolean;
```

Property Value  
boolean

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)