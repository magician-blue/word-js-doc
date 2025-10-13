# Word.Interfaces.BookmarkLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents a single bookmark in a document, selection, or range. The `Bookmark` object is a member of the `Bookmark` collection. The [Word.BookmarkCollection](/en-us/javascript/api/word/word.bookmarkcollection) includes all the bookmarks listed in the Bookmark dialog box (Insert menu).

## Remarks

[ [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

## Properties

- $all: Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
- end: Specifies the ending character position of the bookmark.
- isColumn: Returns `true` if the bookmark is a table column.
- isEmpty: Returns `true` if the bookmark is empty.
- name: Returns the name of the `Bookmark` object.
- range: Returns a `Range` object that represents the portion of the document that's contained in the `Bookmark` object.
- start: Specifies the starting character position of the bookmark.
- storyType: Returns the story type for the bookmark.

## Property Details

### $all

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

#### Property Value
boolean

---

### end

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the ending character position of the bookmark.

```typescript
end?: boolean;
```

#### Property Value
boolean

#### Remarks
[ [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

---

### isColumn

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns `true` if the bookmark is a table column.

```typescript
isColumn?: boolean;
```

#### Property Value
boolean

#### Remarks
[ [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

---

### isEmpty

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns `true` if the bookmark is empty.

```typescript
isEmpty?: boolean;
```

#### Property Value
boolean

#### Remarks
[ [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

---

### name

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the name of the `Bookmark` object.

```typescript
name?: boolean;
```

#### Property Value
boolean

#### Remarks
[ [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

---

### range

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `Range` object that represents the portion of the document that's contained in the `Bookmark` object.

```typescript
range?: Word.Interfaces.RangeLoadOptions;
```

#### Property Value
[Word.Interfaces.RangeLoadOptions](/en-us/javascript/api/word/word.interfaces.rangeloadoptions)

#### Remarks
[ [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

---

### start

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the starting character position of the bookmark.

```typescript
start?: boolean;
```

#### Property Value
boolean

#### Remarks
[ [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

---

### storyType

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the story type for the bookmark.

```typescript
storyType?: boolean;
```

#### Property Value
boolean

#### Remarks
[ [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]