# Word.Interfaces.TrackedChangeCollectionLoadOptions interface

Package: [word](https://learn.microsoft.com/en-us/javascript/api/word)

Contains a collection of [Word.TrackedChange](https://learn.microsoft.com/en-us/javascript/api/word/word.trackedchange) objects.

## Remarks

[ API set: WordApi 1.6 ](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- $all: Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
- author: For EACH ITEM in the collection: Gets the author of the tracked change.
- date: For EACH ITEM in the collection: Gets the date of the tracked change.
- text: For EACH ITEM in the collection: Gets the text of the tracked change.
- type: For EACH ITEM in the collection: Gets the type of the tracked change.

## Property Details

### $all

Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

Property Value: boolean

---

### author

For EACH ITEM in the collection: Gets the author of the tracked change.

```typescript
author?: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApi 1.6 ](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### date

For EACH ITEM in the collection: Gets the date of the tracked change.

```typescript
date?: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApi 1.6 ](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### text

For EACH ITEM in the collection: Gets the text of the tracked change.

```typescript
text?: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApi 1.6 ](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### type

For EACH ITEM in the collection: Gets the type of the tracked change.

```typescript
type?: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApi 1.6 ](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)