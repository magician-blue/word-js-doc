# Word.Interfaces.LineNumberingLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents line numbers in the left margin or to the left of each newspaper-style column.

## Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- $all  
  Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
- countBy  
  Specifies the numeric increment for line numbers.
- distanceFromText  
  Specifies the distance (in points) between the right edge of line numbers and the left edge of the document text.
- isActive  
  Specifies if line numbering is active for the specified document, section, or sections.
- restartMode  
  Specifies the way line numbering runs; that is, whether it starts over at the beginning of a new page or section, or runs continuously.
- startingNumber  
  Specifies the starting line number.

## Property Details

### $all

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

Property Value: boolean

---

### countBy

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the numeric increment for line numbers.

```typescript
countBy?: boolean;
```

Property Value: boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### distanceFromText

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the distance (in points) between the right edge of line numbers and the left edge of the document text.

```typescript
distanceFromText?: boolean;
```

Property Value: boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### isActive

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if line numbering is active for the specified document, section, or sections.

```typescript
isActive?: boolean;
```

Property Value: boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### restartMode

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the way line numbering runs; that is, whether it starts over at the beginning of a new page or section, or runs continuously.

```typescript
restartMode?: boolean;
```

Property Value: boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### startingNumber

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the starting line number.

```typescript
startingNumber?: boolean;
```

Property Value: boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)