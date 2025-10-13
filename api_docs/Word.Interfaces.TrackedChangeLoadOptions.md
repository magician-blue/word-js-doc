# Word.Interfaces.TrackedChangeLoadOptions interface

- Package: https://learn.microsoft.com/en-us/javascript/api/word

Represents a tracked change in a Word document.

## Remarks

[API set: WordApi 1.6](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- $all: Specifying $all for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
- author: Gets the author of the tracked change.
- date: Gets the date of the tracked change.
- text: Gets the text of the tracked change.
- type: Gets the type of the tracked change.

## Property Details

### $all

Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

- Property value: boolean

### author

Gets the author of the tracked change.

```typescript
author?: boolean;
```

- Property value: boolean
- Remarks: [API set: WordApi 1.6](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### date

Gets the date of the tracked change.

```typescript
date?: boolean;
```

- Property value: boolean
- Remarks: [API set: WordApi 1.6](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### text

Gets the text of the tracked change.

```typescript
text?: boolean;
```

- Property value: boolean
- Remarks: [API set: WordApi 1.6](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### type

Gets the type of the tracked change.

```typescript
type?: boolean;
```

- Property value: boolean
- Remarks: [API set: WordApi 1.6](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)