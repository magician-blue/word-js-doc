# Word.Interfaces.TrackedChangeData interface

Package: [word](/en-us/javascript/api/word)

An interface describing the data returned by calling `trackedChange.toJSON()`.

## Properties

- author — Gets the author of the tracked change.
- date — Gets the date of the tracked change.
- text — Gets the text of the tracked change.
- type — Gets the type of the tracked change.

## Property Details

### author

Gets the author of the tracked change.

```typescript
author?: string;
```

Property Value: string

Remarks: [API set: WordApi 1.6](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### date

Gets the date of the tracked change.

```typescript
date?: Date;
```

Property Value: Date

Remarks: [API set: WordApi 1.6](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### text

Gets the text of the tracked change.

```typescript
text?: string;
```

Property Value: string

Remarks: [API set: WordApi 1.6](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### type

Gets the type of the tracked change.

```typescript
type?: Word.TrackedChangeType | "None" | "Added" | "Deleted" | "Formatted";
```

Property Value: [Word.TrackedChangeType](/en-us/javascript/api/word/word.trackedchangetype) | "None" | "Added" | "Deleted" | "Formatted"

Remarks: [API set: WordApi 1.6](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)