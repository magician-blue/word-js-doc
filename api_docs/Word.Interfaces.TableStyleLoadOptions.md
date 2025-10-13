# Word.Interfaces.TableStyleLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Represents the TableStyle object.

## Remarks

[API set: WordApi 1.6](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- $all
  - Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).
- alignment
  - Specifies the table's alignment against the page margin.
- allowBreakAcrossPage
  - Specifies whether lines in tables formatted with a specified style break across pages.
- bottomCellMargin
  - Specifies the amount of space to add between the contents and the bottom borders of the cells.
- cellSpacing
  - Specifies the spacing (in points) between the cells in a table style.
- leftCellMargin
  - Specifies the amount of space to add between the contents and the left borders of the cells.
- rightCellMargin
  - Specifies the amount of space to add between the contents and the right borders of the cells.
- topCellMargin
  - Specifies the amount of space to add between the contents and the top borders of the cells.

## Property Details

### $all

Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).

```typescript
$all?: boolean;
```

Property Value
- boolean

### alignment

Specifies the table's alignment against the page margin.

```typescript
alignment?: boolean;
```

Property Value
- boolean

Remarks
- [API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### allowBreakAcrossPage

Specifies whether lines in tables formatted with a specified style break across pages.

```typescript
allowBreakAcrossPage?: boolean;
```

Property Value
- boolean

Remarks
- [API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### bottomCellMargin

Specifies the amount of space to add between the contents and the bottom borders of the cells.

```typescript
bottomCellMargin?: boolean;
```

Property Value
- boolean

Remarks
- [API set: WordApi 1.6](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### cellSpacing

Specifies the spacing (in points) between the cells in a table style.

```typescript
cellSpacing?: boolean;
```

Property Value
- boolean

Remarks
- [API set: WordApi 1.6](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### leftCellMargin

Specifies the amount of space to add between the contents and the left borders of the cells.

```typescript
leftCellMargin?: boolean;
```

Property Value
- boolean

Remarks
- [API set: WordApi 1.6](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### rightCellMargin

Specifies the amount of space to add between the contents and the right borders of the cells.

```typescript
rightCellMargin?: boolean;
```

Property Value
- boolean

Remarks
- [API set: WordApi 1.6](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### topCellMargin

Specifies the amount of space to add between the contents and the top borders of the cells.

```typescript
topCellMargin?: boolean;
```

Property Value
- boolean

Remarks
- [API set: WordApi 1.6](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)