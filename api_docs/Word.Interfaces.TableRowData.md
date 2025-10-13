# Word.Interfaces.TableRowData interface

Package: [word](/en-us/javascript/api/word)

An interface describing the data returned by calling `tableRow.toJSON()`.

## Properties

- cellCount: Gets the number of cells in the row.
- cells: Gets cells.
- fields: Gets the collection of field objects in the table row.
- font: Gets the font. Use this to get and set font name, size, color, and other properties.
- horizontalAlignment: Specifies the horizontal alignment of every cell in the row. The value can be 'Left', 'Centered', 'Right', or 'Justified'.
- isHeader: Checks whether the row is a header row. To set the number of header rows, use `headerRowCount` on the Table object.
- preferredHeight: Specifies the preferred height of the row in points.
- rowIndex: Gets the index of the row in its parent table.
- shadingColor: Specifies the shading color. Color is specified in "#RRGGBB" format or by using the color name.
- values: Specifies the text values in the row, as a 2D JavaScript array.
- verticalAlignment: Specifies the vertical alignment of the cells in the row. The value can be 'Top', 'Center', or 'Bottom'.

## Property Details

### cellCount

Gets the number of cells in the row.

```typescript
cellCount?: number;
```

Property Value
- number

Remarks  
[API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### cells

Gets cells.

```typescript
cells?: Word.Interfaces.TableCellData[];
```

Property Value
- [Word.Interfaces.TableCellData](/en-us/javascript/api/word/word.interfaces.tablecelldata)[]

Remarks  
[API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### fields

Gets the collection of field objects in the table row.

```typescript
fields?: Word.Interfaces.FieldData[];
```

Property Value
- [Word.Interfaces.FieldData](/en-us/javascript/api/word/word.interfaces.fielddata)[]

Remarks  
[API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### font

Gets the font. Use this to get and set font name, size, color, and other properties.

```typescript
font?: Word.Interfaces.FontData;
```

Property Value
- [Word.Interfaces.FontData](/en-us/javascript/api/word/word.interfaces.fontdata)

Remarks  
[API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### horizontalAlignment

Specifies the horizontal alignment of every cell in the row. The value can be 'Left', 'Centered', 'Right', or 'Justified'.

```typescript
horizontalAlignment?: Word.Alignment | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified";
```

Property Value
- [Word.Alignment](/en-us/javascript/api/word/word.alignment) | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified"

Remarks  
[API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### isHeader

Checks whether the row is a header row. To set the number of header rows, use `headerRowCount` on the Table object.

```typescript
isHeader?: boolean;
```

Property Value
- boolean

Remarks  
[API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### preferredHeight

Specifies the preferred height of the row in points.

```typescript
preferredHeight?: number;
```

Property Value
- number

Remarks  
[API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### rowIndex

Gets the index of the row in its parent table.

```typescript
rowIndex?: number;
```

Property Value
- number

Remarks  
[API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### shadingColor

Specifies the shading color. Color is specified in "#RRGGBB" format or by using the color name.

```typescript
shadingColor?: string;
```

Property Value
- string

Remarks  
[API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### values

Specifies the text values in the row, as a 2D JavaScript array.

```typescript
values?: string[][];
```

Property Value
- string[][]

Remarks  
[API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### verticalAlignment

Specifies the vertical alignment of the cells in the row. The value can be 'Top', 'Center', or 'Bottom'.

```typescript
verticalAlignment?: Word.VerticalAlignment | "Mixed" | "Top" | "Center" | "Bottom";
```

Property Value
- [Word.VerticalAlignment](/en-us/javascript/api/word/word.verticalalignment) | "Mixed" | "Top" | "Center" | "Bottom"

Remarks  
[API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)