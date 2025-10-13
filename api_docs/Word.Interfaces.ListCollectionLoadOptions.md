# Word.Interfaces.ListCollectionLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Contains a collection of [Word.List](/en-us/javascript/api/word/word.list) objects.

## Remarks
[API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties
- $all: Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
- id: For EACH ITEM in the collection: Gets the list's id.
- levelExistences: For EACH ITEM in the collection: Checks whether each of the 9 levels exists in the list. A true value indicates the level exists, which means there is at least one list item at that level.
- levelTypes: For EACH ITEM in the collection: Gets all 9 level types in the list. Each type can be 'Bullet', 'Number', or 'Picture'.

## Property Details

### $all
Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

Property Value
- boolean

---

### id
For EACH ITEM in the collection: Gets the list's id.

```typescript
id?: boolean;
```

Property Value
- boolean

Remarks
- [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### levelExistences
For EACH ITEM in the collection: Checks whether each of the 9 levels exists in the list. A true value indicates the level exists, which means there is at least one list item at that level.

```typescript
levelExistences?: boolean;
```

Property Value
- boolean

Remarks
- [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### levelTypes
For EACH ITEM in the collection: Gets all 9 level types in the list. Each type can be 'Bullet', 'Number', or 'Picture'.

```typescript
levelTypes?: boolean;
```

Property Value
- boolean

Remarks
- [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)