# Word.Interfaces.LineNumberingData interface

Package: [word](/en-us/javascript/api/word)

An interface describing the data returned by calling `lineNumbering.toJSON()`.

## Properties

- countBy: Specifies the numeric increment for line numbers.
- distanceFromText: Specifies the distance (in points) between the right edge of line numbers and the left edge of the document text.
- isActive: Specifies if line numbering is active for the specified document, section, or sections.
- restartMode: Specifies the way line numbering runs; that is, whether it starts over at the beginning of a new page or section, or runs continuously.
- startingNumber: Specifies the starting line number.

## Property Details

### countBy

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the numeric increment for line numbers.

```typescript
countBy?: number;
```

#### Property Value
number

#### Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### distanceFromText

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the distance (in points) between the right edge of line numbers and the left edge of the document text.

```typescript
distanceFromText?: number;
```

#### Property Value
number

#### Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### isActive

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if line numbering is active for the specified document, section, or sections.

```typescript
isActive?: boolean;
```

#### Property Value
boolean

#### Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### restartMode

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the way line numbering runs; that is, whether it starts over at the beginning of a new page or section, or runs continuously.

```typescript
restartMode?: Word.NumberingRule | "RestartContinuous" | "RestartSection" | "RestartPage";
```

#### Property Value
[Word.NumberingRule](/en-us/javascript/api/word/word.numberingrule) | "RestartContinuous" | "RestartSection" | "RestartPage"

#### Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### startingNumber

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the starting line number.

```typescript
startingNumber?: number;
```

#### Property Value
number

#### Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]