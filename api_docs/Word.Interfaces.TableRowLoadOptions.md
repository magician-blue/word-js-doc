# Word.Interfaces.TableRowLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Represents a row in a Word document.

## Remarks

[API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- $all: Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).
- cellCount: Gets the number of cells in the row.
- font: Gets the font. Use this to get and set font name, size, color, and other properties.
- horizontalAlignment: Specifies the horizontal alignment of every cell in the row. The value can be 'Left', 'Centered', 'Right', or 'Justified'.
- isHeader: Checks whether the row is a header row. To set the number of header rows, use headerRowCount on the Table object.
- parentTable: Gets parent table.
- preferredHeight: Specifies the preferred height of the row in points.
- rowIndex: Gets the index of the row in its parent table.
- shadingColor: Specifies the shading color. Color is specified in "#RRGGBB" format or by using the color name.
- values: Specifies the text values in the row, as a 2D JavaScript array.
- verticalAlignment: Specifies the vertical alignment of the cells in the row. The value can be 'Top', 'Center', or 'Bottom'.

## Property Details

### $all

Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).

```typescript
$all?: boolean;
```

- Property Value: boolean

---

### cellCount

Gets the number of cells in the row.

```typescript
cellCount?: boolean;
```

- Property Value: boolean
- Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### font

Gets the font. Use this to get and set font name, size, color, and other properties.

```typescript
font?: Word.Interfaces.FontLoadOptions;
```

- Property Value: [Word.Interfaces.FontLoadOptions](/en-us/javascript/api/word/word.interfaces.fontloadoptions)
- Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### horizontalAlignment

Specifies the horizontal alignment of every cell in the row. The value can be 'Left', 'Centered', 'Right', or 'Justified'.

```typescript
horizontalAlignment?: boolean;
```

- Property Value: boolean
- Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### isHeader

Checks whether the row is a header row. To set the number of header rows, use headerRowCount on the Table object.

```typescript
isHeader?: boolean;
```

- Property Value: boolean
- Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### parentTable

Gets parent table.

```typescript
parentTable?: Word.Interfaces.TableLoadOptions;
```

- Property Value: [Word.Interfaces.TableLoadOptions](/en-us/javascript/api/word/word.interfaces.tableloadoptions)
- Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### preferredHeight

Specifies the preferred height of the row in points.

```typescript
preferredHeight?: boolean;
```

- Property Value: boolean
- Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### rowIndex

Gets the index of the row in its parent table.

```typescript
rowIndex?: boolean;
```

- Property Value: boolean
- Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### shadingColor

Specifies the shading color. Color is specified in "#RRGGBB" format or by using the color name.

```typescript
shadingColor?: boolean;
```

- Property Value: boolean
- Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### values

Specifies the text values in the row, as a 2D JavaScript array.

```typescript
values?: boolean;
```

- Property Value: boolean
- Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### verticalAlignment

Specifies the vertical alignment of the cells in the row. The value can be 'Top', 'Center', or 'Bottom'.

```typescript
verticalAlignment?: boolean;
```

- Property Value: boolean
- Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)