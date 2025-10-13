# Word.Interfaces.ListFormatUpdateData interface

Package: [word](/en-us/javascript/api/word)

An interface for updating data on the ListFormat object, for use in listFormat.set({ ... }).

## Properties

- listLevelNumber  
  Specifies the list level number for the first paragraph for the ListFormat object.
- listTemplate  
  Gets the list template associated with the ListFormat object.

## Property Details

### listLevelNumber

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the list level number for the first paragraph for the ListFormat object.

```typescript
listLevelNumber?: number;
```

- Property Value: number

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### listTemplate

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the list template associated with the ListFormat object.

```typescript
listTemplate?: Word.Interfaces.ListTemplateUpdateData;
```

- Property Value: [Word.Interfaces.ListTemplateUpdateData](/en-us/javascript/api/word/word.interfaces.listtemplateupdatedata)

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)