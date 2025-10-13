# Word.Interfaces.ShapeGroupLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Represents a shape group in the document. To get the corresponding Shape object, use `ShapeGroup.shape`.

## Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties
- `$all` — Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
- `id` — Gets an integer that represents the shape group identifier.
- `shape` — Gets the Shape object associated with the group.

## Property Details

### $all
Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

Property value: boolean

---

### id
Gets an integer that represents the shape group identifier.

```typescript
id?: boolean;
```

Property value: boolean

Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### shape
Gets the Shape object associated with the group.

```typescript
shape?: Word.Interfaces.ShapeLoadOptions;
```

Property value: [Word.Interfaces.ShapeLoadOptions](/en-us/javascript/api/word/word.interfaces.shapeloadoptions)

Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)