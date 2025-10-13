# Word.Interfaces.BreakData interface

Package: [word](/en-us/javascript/api/word)

An interface describing the data returned by calling `break.toJSON()`.

## Properties

- pageIndex — Returns the page number on which the break occurs.
- range — Returns a `Range` object that represents the portion of the document that's contained in the break.

## Property Details

### pageIndex

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the page number on which the break occurs.

```typescript
pageIndex?: number;
```

#### Property Value
number

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### range

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `Range` object that represents the portion of the document that's contained in the break.

```typescript
range?: Word.Interfaces.RangeData;
```

#### Property Value
[Word.Interfaces.RangeData](/en-us/javascript/api/word/word.interfaces.rangedata)

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)