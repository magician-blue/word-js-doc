# Word.Interfaces.ListData interface

Package: [word](https://learn.microsoft.com/en-us/javascript/api/word)

An interface describing the data returned by calling `list.toJSON()`.

## Properties

- `id`: Gets the list's id.
- `levelExistences`: Checks whether each of the 9 levels exists in the list. A true value indicates the level exists, which means there is at least one list item at that level.
- `levelTypes`: Gets all 9 level types in the list. Each type can be 'Bullet', 'Number', or 'Picture'.
- `paragraphs`: Gets paragraphs in the list.

## Property Details

### id

Gets the list's id.

```typescript
id?: number;
```

- Property Value: number

Remarks: [API set: WordApi 1.3](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### levelExistences

Checks whether each of the 9 levels exists in the list. A true value indicates the level exists, which means there is at least one list item at that level.

```typescript
levelExistences?: boolean[];
```

- Property Value: boolean[]

Remarks: [API set: WordApi 1.3](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### levelTypes

Gets all 9 level types in the list. Each type can be 'Bullet', 'Number', or 'Picture'.

```typescript
levelTypes?: Word.ListLevelType[];
```

- Property Value: [Word.ListLevelType](https://learn.microsoft.com/en-us/javascript/api/word/word.listleveltype)[]

Remarks: [API set: WordApi 1.3](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### paragraphs

Gets paragraphs in the list.

```typescript
paragraphs?: Word.Interfaces.ParagraphData[];
```

- Property Value: [Word.Interfaces.ParagraphData](https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.paragraphdata)[]

Remarks: [API set: WordApi 1.3](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)