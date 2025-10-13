# Word.Interfaces.BookmarkData interface

Package: [word](/en-us/javascript/api/word)

An interface describing the data returned by calling bookmark.toJSON().

## Properties

- end — Specifies the ending character position of the bookmark.
- isColumn — Returns true if the bookmark is a table column.
- isEmpty — Returns true if the bookmark is empty.
- name — Returns the name of the Bookmark object.
- range — Returns a Range object that represents the portion of the document that's contained in the Bookmark object.
- start — Specifies the starting character position of the bookmark.
- storyType — Returns the story type for the bookmark.

## Property Details

### end

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the ending character position of the bookmark.

```typescript
end?: number;
```

Property Value
- number

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### isColumn

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns true if the bookmark is a table column.

```typescript
isColumn?: boolean;
```

Property Value
- boolean

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### isEmpty

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns true if the bookmark is empty.

```typescript
isEmpty?: boolean;
```

Property Value
- boolean

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### name

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the name of the Bookmark object.

```typescript
name?: string;
```

Property Value
- string

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### range

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a Range object that represents the portion of the document that's contained in the Bookmark object.

```typescript
range?: Word.Interfaces.RangeData;
```

Property Value
- [Word.Interfaces.RangeData](/en-us/javascript/api/word/word.interfaces.rangedata)

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### start

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the starting character position of the bookmark.

```typescript
start?: number;
```

Property Value
- number

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### storyType

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the story type for the bookmark.

```typescript
storyType?: Word.StoryType | "MainText" | "Footnotes" | "Endnotes" | "Comments" | "TextFrame" | "EvenPagesHeader" | "PrimaryHeader" | "EvenPagesFooter" | "PrimaryFooter" | "FirstPageHeader" | "FirstPageFooter" | "FootnoteSeparator" | "FootnoteContinuationSeparator" | "FootnoteContinuationNotice" | "EndnoteSeparator" | "EndnoteContinuationSeparator" | "EndnoteContinuationNotice";
```

Property Value
- [Word.StoryType](/en-us/javascript/api/word/word.storytype) | "MainText" | "Footnotes" | "Endnotes" | "Comments" | "TextFrame" | "EvenPagesHeader" | "PrimaryHeader" | "EvenPagesFooter" | "PrimaryFooter" | "FirstPageHeader" | "FirstPageFooter" | "FootnoteSeparator" | "FootnoteContinuationSeparator" | "FootnoteContinuationNotice" | "EndnoteSeparator" | "EndnoteContinuationSeparator" | "EndnoteContinuationNotice"

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)