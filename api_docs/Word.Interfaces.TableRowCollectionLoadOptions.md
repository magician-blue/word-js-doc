# Word.Interfaces.TableRowCollectionLoadOptions interface

Package: word

Contains the collection of the document's TableRow objects.

## Remarks
- API set: WordApi 1.3

## Properties
- $all — Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).
- cellCount — For EACH ITEM in the collection: Gets the number of cells in the row.
- font — For EACH ITEM in the collection: Gets the font. Use this to get and set font name, size, color, and other properties.
- horizontalAlignment — For EACH ITEM in the collection: Specifies the horizontal alignment of every cell in the row. The value can be 'Left', 'Centered', 'Right', or 'Justified'.
- isHeader — For EACH ITEM in the collection: Checks whether the row is a header row. To set the number of header rows, use headerRowCount on the Table object.
- parentTable — For EACH ITEM in the collection: Gets parent table.
- preferredHeight — For EACH ITEM in the collection: Specifies the preferred height of the row in points.
- rowIndex — For EACH ITEM in the collection: Gets the index of the row in its parent table.
- shadingColor — For EACH ITEM in the collection: Specifies the shading color. Color is specified in "#RRGGBB" format or by using the color name.
- values — For EACH ITEM in the collection: Specifies the text values in the row, as a 2D JavaScript array.
- verticalAlignment — For EACH ITEM in the collection: Specifies the vertical alignment of the cells in the row. The value can be 'Top', 'Center', or 'Bottom'.

## Property Details

### $all
Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).

```typescript
$all?: boolean;
```

Property value: boolean

---

### cellCount
For EACH ITEM in the collection: Gets the number of cells in the row.

```typescript
cellCount?: boolean;
```

Property value: boolean

Remarks:
- API set: WordApi 1.3

---

### font
For EACH ITEM in the collection: Gets the font. Use this to get and set font name, size, color, and other properties.

```typescript
font?: Word.Interfaces.FontLoadOptions;
```

Property value: Word.Interfaces.FontLoadOptions

Remarks:
- API set: WordApi 1.3

---

### horizontalAlignment
For EACH ITEM in the collection: Specifies the horizontal alignment of every cell in the row. The value can be 'Left', 'Centered', 'Right', or 'Justified'.

```typescript
horizontalAlignment?: boolean;
```

Property value: boolean

Remarks:
- API set: WordApi 1.3

---

### isHeader
For EACH ITEM in the collection: Checks whether the row is a header row. To set the number of header rows, use headerRowCount on the Table object.

```typescript
isHeader?: boolean;
```

Property value: boolean

Remarks:
- API set: WordApi 1.3

---

### parentTable
For EACH ITEM in the collection: Gets parent table.

```typescript
parentTable?: Word.Interfaces.TableLoadOptions;
```

Property value: Word.Interfaces.TableLoadOptions

Remarks:
- API set: WordApi 1.3

---

### preferredHeight
For EACH ITEM in the collection: Specifies the preferred height of the row in points.

```typescript
preferredHeight?: boolean;
```

Property value: boolean

Remarks:
- API set: WordApi 1.3

---

### rowIndex
For EACH ITEM in the collection: Gets the index of the row in its parent table.

```typescript
rowIndex?: boolean;
```

Property value: boolean

Remarks:
- API set: WordApi 1.3

---

### shadingColor
For EACH ITEM in the collection: Specifies the shading color. Color is specified in "#RRGGBB" format or by using the color name.

```typescript
shadingColor?: boolean;
```

Property value: boolean

Remarks:
- API set: WordApi 1.3

---

### values
For EACH ITEM in the collection: Specifies the text values in the row, as a 2D JavaScript array.

```typescript
values?: boolean;
```

Property value: boolean

Remarks:
- API set: WordApi 1.3

---

### verticalAlignment
For EACH ITEM in the collection: Specifies the vertical alignment of the cells in the row. The value can be 'Top', 'Center', or 'Bottom'.

```typescript
verticalAlignment?: boolean;
```

Property value: boolean

Remarks:
- API set: WordApi 1.3