# Word.CustomXmlAddValidationErrorOptions interface

Package: [word](/en-us/javascript/api/word)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

The options that define the descriptive error text and the state of `clearedOnUpdate`.

## Remarks

[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- clearedOnUpdate  
  If provided, specifies whether the error is to be cleared from the [Word.CustomXmlValidationErrorCollection](/en-us/javascript/api/word/word.customxmlvalidationerrorcollection) when the XML is corrected and updated.

- errorText  
  If provided, specifies the descriptive error text.

## Property Details

### clearedOnUpdate

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

If provided, specifies whether the error is to be cleared from the [Word.CustomXmlValidationErrorCollection](/en-us/javascript/api/word/word.customxmlvalidationerrorcollection) when the XML is corrected and updated.

```typescript
clearedOnUpdate?: boolean;
```

Property Value  
boolean

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### errorText

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

If provided, specifies the descriptive error text.

```typescript
errorText?: string;
```

Property Value  
string

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)