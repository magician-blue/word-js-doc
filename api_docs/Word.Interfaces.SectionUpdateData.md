# Word.Interfaces.SectionUpdateData interface

Package: https://learn.microsoft.com/en-us/javascript/api/word

An interface for updating data on the Section object, for use in section.set({ ... }).

## Properties

- body
  - Gets the body object of the section. This doesn't include the header/footer and other section metadata.
- pageSetup
  - Returns a PageSetup object that's associated with the section.
- protectedForForms
  - Specifies if the section is protected for forms.

## Property Details

### body

Gets the body object of the section. This doesn't include the header/footer and other section metadata.

```typescript
body?: Word.Interfaces.BodyUpdateData;
```

Property Value:
- Word.Interfaces.BodyUpdateData: https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.bodyupdatedata

Remarks:
- [API set: WordApi 1.1](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### pageSetup

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a PageSetup object that's associated with the section.

```typescript
pageSetup?: Word.Interfaces.PageSetupUpdateData;
```

Property Value:
- Word.Interfaces.PageSetupUpdateData: https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.pagesetupupdatedata

Remarks:
- [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### protectedForForms

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the section is protected for forms.

```typescript
protectedForForms?: boolean;
```

Property Value:
- boolean

Remarks:
- [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)