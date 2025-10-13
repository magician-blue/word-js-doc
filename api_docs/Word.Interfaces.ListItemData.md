# Word.Interfaces.ListItemData interface

Package: [word](/en-us/javascript/api/word)

An interface describing the data returned by calling listItem.toJSON().

## Properties

- [level](#level)
  - Specifies the level of the item in the list.
- [listString](#liststring)
  - Gets the list item bullet, number, or picture as a string.
- [siblingIndex](#siblingindex)
  - Gets the list item order number in relation to its siblings.

## Property details

### level

Specifies the level of the item in the list.

```typescript
level?: number;
```

#### Property Value
number

#### Remarks
[API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### listString

Gets the list item bullet, number, or picture as a string.

```typescript
listString?: string;
```

#### Property Value
string

#### Remarks
[API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### siblingIndex

Gets the list item order number in relation to its siblings.

```typescript
siblingIndex?: number;
```

#### Property Value
number

#### Remarks
[API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)