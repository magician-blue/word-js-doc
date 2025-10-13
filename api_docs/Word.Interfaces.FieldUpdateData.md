# Word.Interfaces.FieldUpdateData interface

Package: [word](/en-us/javascript/api/word)

An interface for updating data on the `Field` object, for use in `field.set({ ... })`.

## Properties

- code — Specifies the field's code instruction.
- data — Specifies data in an "Addin" field. If the field isn't an "Addin" field, it is `null` and it will throw a general exception when code attempts to set it.
- locked — Specifies whether the field is locked. `true` if the field is locked, `false` otherwise.
- result — Gets the field's result data.
- showCodes — Specifies whether the field codes are displayed for the specified field. `true` if the field codes are displayed, `false` otherwise.

## Property Details

### code

Specifies the field's code instruction.

```typescript
code?: string;
```

- Property Value: string
- Remarks:
  - [API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
  - Note: The ability to set the code was introduced in WordApi 1.5.

### data

Specifies data in an "Addin" field. If the field isn't an "Addin" field, it is `null` and it will throw a general exception when code attempts to set it.

```typescript
data?: string;
```

- Property Value: string
- Remarks:
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### locked

Specifies whether the field is locked. `true` if the field is locked, `false` otherwise.

```typescript
locked?: boolean;
```

- Property Value: boolean
- Remarks:
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### result

Gets the field's result data.

```typescript
result?: Word.Interfaces.RangeUpdateData;
```

- Property Value: [Word.Interfaces.RangeUpdateData](/en-us/javascript/api/word/word.interfaces.rangeupdatedata)
- Remarks:
  - [API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### showCodes

Specifies whether the field codes are displayed for the specified field. `true` if the field codes are displayed, `false` otherwise.

```typescript
showCodes?: boolean;
```

- Property Value: boolean
- Remarks:
  - [API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)