# Word.Interfaces.TableStyleUpdateData interface

Package: [word](/en-us/javascript/api/word)

An interface for updating data on the TableStyle object, for use in tableStyle.set({ ... }).

## Properties

- `alignment` — Specifies the table's alignment against the page margin.
- `allowBreakAcrossPage` — Specifies whether lines in tables formatted with a specified style break across pages.
- `bottomCellMargin` — Specifies the amount of space to add between the contents and the bottom borders of the cells.
- `cellSpacing` — Specifies the spacing (in points) between the cells in a table style.
- `leftCellMargin` — Specifies the amount of space to add between the contents and the left borders of the cells.
- `rightCellMargin` — Specifies the amount of space to add between the contents and the right borders of the cells.
- `topCellMargin` — Specifies the amount of space to add between the contents and the top borders of the cells.

## Property Details

### alignment

Specifies the table's alignment against the page margin.

```typescript
alignment?: Word.Alignment | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified";
```

Property Value:
[Word.Alignment](/en-us/javascript/api/word/word.alignment) | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified"

Remarks:
[API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### allowBreakAcrossPage

Specifies whether lines in tables formatted with a specified style break across pages.

```typescript
allowBreakAcrossPage?: boolean;
```

Property Value:
boolean

Remarks:
[API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### bottomCellMargin

Specifies the amount of space to add between the contents and the bottom borders of the cells.

```typescript
bottomCellMargin?: number;
```

Property Value:
number

Remarks:
[API set: WordApi 1.6](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### cellSpacing

Specifies the spacing (in points) between the cells in a table style.

```typescript
cellSpacing?: number;
```

Property Value:
number

Remarks:
[API set: WordApi 1.6](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### leftCellMargin

Specifies the amount of space to add between the contents and the left borders of the cells.

```typescript
leftCellMargin?: number;
```

Property Value:
number

Remarks:
[API set: WordApi 1.6](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### rightCellMargin

Specifies the amount of space to add between the contents and the right borders of the cells.

```typescript
rightCellMargin?: number;
```

Property Value:
number

Remarks:
[API set: WordApi 1.6](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### topCellMargin

Specifies the amount of space to add between the contents and the top borders of the cells.

```typescript
topCellMargin?: number;
```

Property Value:
number

Remarks:
[API set: WordApi 1.6](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)