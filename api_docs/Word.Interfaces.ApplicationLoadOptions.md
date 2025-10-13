# Word.Interfaces.ApplicationLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Represents the application object.

## Remarks

[ API set: WordApi 1.3 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- $all  
  Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).

- bibliography  
  Returns a Bibliography object that represents the bibliography reference sources stored in Microsoft Word.

- checkLanguage  
  Specifies if Microsoft Word automatically detects the language you are using as you type.

- language  
  Gets a LanguageId value that represents the language selected for the Microsoft Word user interface.

## Property Details

### $all

Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).

```typescript
$all?: boolean;
```

- Property Value: boolean

---

### bibliography

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a Bibliography object that represents the bibliography reference sources stored in Microsoft Word.

```typescript
bibliography?: Word.Interfaces.BibliographyLoadOptions;
```

- Property Value: [Word.Interfaces.BibliographyLoadOptions](/en-us/javascript/api/word/word.interfaces.bibliographyloadoptions)

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### checkLanguage

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if Microsoft Word automatically detects the language you are using as you type.

```typescript
checkLanguage?: boolean;
```

- Property Value: boolean

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### language

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets a LanguageId value that represents the language selected for the Microsoft Word user interface.

```typescript
language?: boolean;
```

- Property Value: boolean

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)