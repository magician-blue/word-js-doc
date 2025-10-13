# Word.TextColumnAddOptions interface

Package: [word](/en-us/javascript/api/word)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents options for a new text column in a document or section of a document.

## Remarks

[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- isEvenlySpaced  
  If provided, specifies whether to evenly space all the text columns in the document. The default value is `true`.

- spacing  
  If provided, specifies the spacing between the text columns in the document, in points. The default value is -1, which means Word will automatically determine the width based on the number of columns and page size.

- width  
  If provided, specifies the width of the new text column in the document, in points. The default value is -1, which means Word will automatically determine the width based on the number of columns and page size.

## Property Details

### isEvenlySpaced

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

If provided, specifies whether to evenly space all the text columns in the document. The default value is `true`.

```typescript
isEvenlySpaced?: boolean;
```

Property Value  
boolean

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### spacing

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

If provided, specifies the spacing between the text columns in the document, in points. The default value is -1, which means Word will automatically determine the width based on the number of columns and page size.

```typescript
spacing?: number;
```

Property Value  
number

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### width

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

If provided, specifies the width of the new text column in the document, in points. The default value is -1, which means Word will automatically determine the width based on the number of columns and page size.

```typescript
width?: number;
```

Property Value  
number

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)