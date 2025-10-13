# Word.Interfaces.BookmarkCollectionLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

A collection of [Word.Bookmark](/en-us/javascript/api/word/word.bookmark) objects that represent the bookmarks in the specified selection, range, or document.

## Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- $all
  - Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
- end
  - For EACH ITEM in the collection: Specifies the ending character position of the bookmark.
- isColumn
  - For EACH ITEM in the collection: Returns `true` if the bookmark is a table column.
- isEmpty
  - For EACH ITEM in the collection: Returns `true` if the bookmark is empty.
- name
  - For EACH ITEM in the collection: Returns the name of the `Bookmark` object.
- range
  - For EACH ITEM in the collection: Returns a `Range` object that represents the portion of the document that's contained in the `Bookmark` object.
- start
  - For EACH ITEM in the collection: Specifies the starting character position of the bookmark.
- storyType
  - For EACH ITEM in the collection: Returns the story type for the bookmark.

## Property Details

### $all

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```
$all?: boolean;
```

Property Value
- boolean

### end

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies the ending character position of the bookmark.

```
end?: boolean;
```

Property Value
- boolean

Remarks
- [ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### isColumn

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Returns `true` if the bookmark is a table column.

```
isColumn?: boolean;
```

Property Value
- boolean

Remarks
- [ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### isEmpty

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Returns `true` if the bookmark is empty.

```
isEmpty?: boolean;
```

Property Value
- boolean

Remarks
- [ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### name

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Returns the name of the `Bookmark` object.

```
name?: boolean;
```

Property Value
- boolean

Remarks
- [ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### range

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Returns a `Range` object that represents the portion of the document that's contained in the `Bookmark` object.

```
range?: Word.Interfaces.RangeLoadOptions;
```

Property Value
- [Word.Interfaces.RangeLoadOptions](/en-us/javascript/api/word/word.interfaces.rangeloadoptions)

Remarks
- [ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### start

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies the starting character position of the bookmark.

```
start?: boolean;
```

Property Value
- boolean

Remarks
- [ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### storyType

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Returns the story type for the bookmark.

```
storyType?: boolean;
```

Property Value
- boolean

Remarks
- [ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)