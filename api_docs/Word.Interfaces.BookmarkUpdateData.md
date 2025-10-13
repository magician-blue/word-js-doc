# Word.Interfaces.BookmarkUpdateData interface

Package: https://learn.microsoft.com/en-us/javascript/api/word

An interface for updating data on the Bookmark object, for use in `bookmark.set({ ... })`.

## Properties

- end — Specifies the ending character position of the bookmark.
- range — Returns a Range object that represents the portion of the document that's contained in the Bookmark object.
- start — Specifies the starting character position of the bookmark.

## Property Details

### end

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the ending character position of the bookmark.

```typescript
end?: number;
```

Property Value: number

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### range

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `Range` object that represents the portion of the document that's contained in the `Bookmark` object.

```typescript
range?: Word.Interfaces.RangeUpdateData;
```

Property Value: https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.rangeupdatedata

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### start

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the starting character position of the bookmark.

```typescript
start?: number;
```

Property Value: number

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)