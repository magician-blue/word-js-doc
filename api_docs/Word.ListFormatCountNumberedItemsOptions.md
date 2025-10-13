# Word.ListFormatCountNumberedItemsOptions interface

Package: [word](https://learn.microsoft.com/en-us/javascript/api/word)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents options for counting numbered items in a range.

## Remarks

[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- level  
  If provided, specifies the level to count. The default value is 1.

- numberType  
  If provided, specifies the type of number to count. The default value is `Word.NumberType.paragraph`.

## Property Details

### level

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

If provided, specifies the level to count. The default value is 1.

```typescript
level?: number;
```

- Type: number

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### numberType

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

If provided, specifies the type of number to count. The default value is `Word.NumberType.paragraph`.

```typescript
numberType?: Word.NumberType | "Paragraph" | "ListNum" | "AllNumbers";
```

- Type: [Word.NumberType](https://learn.microsoft.com/en-us/javascript/api/word/word.numbertype) | "Paragraph" | "ListNum" | "AllNumbers"

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)