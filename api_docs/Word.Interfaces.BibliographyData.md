# Word.Interfaces.BibliographyData interface

Package: [word](/en-us/javascript/api/word)

An interface describing the data returned by calling `bibliography.toJSON()`.

## Properties

- bibliographyStyle
  - Specifies the name of the active style to use for the bibliography.

- sources
  - Returns a `SourceCollection` object that represents all the sources contained in the bibliography.

## Property Details

### bibliographyStyle

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the name of the active style to use for the bibliography.

```typescript
bibliographyStyle?: string;
```

Property Value

- string

Remarks

- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### sources

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `SourceCollection` object that represents all the sources contained in the bibliography.

```typescript
sources?: Word.Interfaces.SourceData[];
```

Property Value

- [Word.Interfaces.SourceData](/en-us/javascript/api/word/word.interfaces.sourcedata)[]

Remarks

- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)