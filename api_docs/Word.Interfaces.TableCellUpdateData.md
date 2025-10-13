# Word.Interfaces.TableCellUpdateData interface

Package: [word](/en-us/javascript/api/word)

An interface for updating data on the TableCell object, for use in tableCell.set({ ... }).

## Properties

- body: Gets the body object of the cell.
- columnWidth: Specifies the width of the cell's column in points. This is applicable to uniform tables.
- horizontalAlignment: Specifies the horizontal alignment of the cell. The value can be 'Left', 'Centered', 'Right', or 'Justified'.
- shadingColor: Specifies the shading color of the cell. Color is specified in "#RRGGBB" format or by using the color name.
- value: Specifies the text of the cell.
- verticalAlignment: Specifies the vertical alignment of the cell. The value can be 'Top', 'Center', or 'Bottom'.

## Property Details

### body

Gets the body object of the cell.

```typescript
body?: Word.Interfaces.BodyUpdateData;
```

#### Property Value
[Word.Interfaces.BodyUpdateData](/en-us/javascript/api/word/word.interfaces.bodyupdatedata)

#### Remarks
[ API set: WordApi 1.3 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### columnWidth

Specifies the width of the cell's column in points. This is applicable to uniform tables.

```typescript
columnWidth?: number;
```

#### Property Value
number

#### Remarks
[ API set: WordApi 1.3 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### horizontalAlignment

Specifies the horizontal alignment of the cell. The value can be 'Left', 'Centered', 'Right', or 'Justified'.

```typescript
horizontalAlignment?: Word.Alignment | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified";
```

#### Property Value
[Word.Alignment](/en-us/javascript/api/word/word.alignment) | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified"

#### Remarks
[ API set: WordApi 1.3 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### shadingColor

Specifies the shading color of the cell. Color is specified in "#RRGGBB" format or by using the color name.

```typescript
shadingColor?: string;
```

#### Property Value
string

#### Remarks
[ API set: WordApi 1.3 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### value

Specifies the text of the cell.

```typescript
value?: string;
```

#### Property Value
string

#### Remarks
[ API set: WordApi 1.3 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### verticalAlignment

Specifies the vertical alignment of the cell. The value can be 'Top', 'Center', or 'Bottom'.

```typescript
verticalAlignment?: Word.VerticalAlignment | "Mixed" | "Top" | "Center" | "Bottom";
```

#### Property Value
[Word.VerticalAlignment](/en-us/javascript/api/word/word.verticalalignment) | "Mixed" | "Top" | "Center" | "Bottom"

#### Remarks
[ API set: WordApi 1.3 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)