# Word.Interfaces.BreakCollectionLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Contains a collection of [Word.Break](/en-us/javascript/api/word/word.break) objects.

## Remarks
[ [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

## Properties

- $all — Specifying $all for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
- pageIndex — For EACH ITEM in the collection: Returns the page number on which the break occurs.
- range — For EACH ITEM in the collection: Returns a `Range` object that represents the portion of the document that's contained in the break.

## Property Details

### $all

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

Property Value
- boolean

### pageIndex

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Returns the page number on which the break occurs.

```typescript
pageIndex?: boolean;
```

Property Value
- boolean

Remarks
[ [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

### range

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Returns a `Range` object that represents the portion of the document that's contained in the break.

```typescript
range?: Word.Interfaces.RangeLoadOptions;
```

Property Value
- [Word.Interfaces.RangeLoadOptions](/en-us/javascript/api/word/word.interfaces.rangeloadoptions)

Remarks
[ [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]