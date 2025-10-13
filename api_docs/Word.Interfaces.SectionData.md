# Word.Interfaces.SectionData interface

Package: [word](/en-us/javascript/api/word)

An interface describing the data returned by calling `section.toJSON()`.

## Properties

- body  
  Gets the body object of the section. This doesn't include the header/footer and other section metadata.
- borders  
  Returns a `BorderUniversalCollection` object that represents all the borders in the section.
- pageSetup  
  Returns a `PageSetup` object that's associated with the section.
- protectedForForms  
  Specifies if the section is protected for forms.

## Property Details

### body

Gets the body object of the section. This doesn't include the header/footer and other section metadata.

```typescript
body?: Word.Interfaces.BodyData;
```

- Property Value: [Word.Interfaces.BodyData](/en-us/javascript/api/word/word.interfaces.bodydata)
- Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### borders

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `BorderUniversalCollection` object that represents all the borders in the section.

```typescript
borders?: Word.Interfaces.BorderUniversalData[];
```

- Property Value: [Word.Interfaces.BorderUniversalData](/en-us/javascript/api/word/word.interfaces.borderuniversaldata)[]
- Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### pageSetup

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `PageSetup` object that's associated with the section.

```typescript
pageSetup?: Word.Interfaces.PageSetupData;
```

- Property Value: [Word.Interfaces.PageSetupData](/en-us/javascript/api/word/word.interfaces.pagesetupdata)
- Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### protectedForForms

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the section is protected for forms.

```typescript
protectedForForms?: boolean;
```

- Property Value: boolean
- Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)