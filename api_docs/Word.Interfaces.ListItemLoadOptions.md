# Word.Interfaces.ListItemLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Represents the paragraph list item format.

## Remarks

[API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- $all: Specifying "$all" for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).
- level: Specifies the level of the item in the list.
- listString: Gets the list item bullet, number, or picture as a string.
- siblingIndex: Gets the list item order number in relation to its siblings.

## Property Details

### $all

Specifying "$all" for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).

```typescript
$all?: boolean;
```

Property Value: boolean

### level

Specifies the level of the item in the list.

```typescript
level?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### listString

Gets the list item bullet, number, or picture as a string.

```typescript
listString?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### siblingIndex

Gets the list item order number in relation to its siblings.

```typescript
siblingIndex?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)