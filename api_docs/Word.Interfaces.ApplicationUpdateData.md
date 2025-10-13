# Word.Interfaces.ApplicationUpdateData interface

Package: [word](/en-us/javascript/api/word)

An interface for updating data on the `Application` object, for use in `application.set({ ... })`.

## Properties

- [bibliography](#bibliography)
  - Returns a `Bibliography` object that represents the bibliography reference sources stored in Microsoft Word.
- [checkLanguage](#checklanguage)
  - Specifies if Microsoft Word automatically detects the language you are using as you type.

## Property Details

### bibliography

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `Bibliography` object that represents the bibliography reference sources stored in Microsoft Word.

```typescript
bibliography?: Word.Interfaces.BibliographyUpdateData;
```

Property Value: [Word.Interfaces.BibliographyUpdateData](/en-us/javascript/api/word/word.interfaces.bibliographyupdatedata)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### checkLanguage

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if Microsoft Word automatically detects the language you are using as you type.

```typescript
checkLanguage?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)