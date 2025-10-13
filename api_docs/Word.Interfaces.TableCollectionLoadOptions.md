# Word.Interfaces.TableCollectionLoadOptions interface

Package: word

Contains the collection of the document's Table objects.

## Remarks
[API set: WordApi 1.3]

## Properties

- $all — Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).
- alignment — For EACH ITEM in the collection: Specifies the alignment of the table against the page column. The value can be 'Left', 'Centered', or 'Right'.
- font — For EACH ITEM in the collection: Gets the font. Use this to get and set font name, size, color, and other properties.
- headerRowCount — For EACH ITEM in the collection: Specifies the number of header rows.
- horizontalAlignment — For EACH ITEM in the collection: Specifies the horizontal alignment of every cell in the table. The value can be 'Left', 'Centered', 'Right', or 'Justified'.
- isUniform — For EACH ITEM in the collection: Indicates whether all of the table rows are uniform.
- nestingLevel — For EACH ITEM in the collection: Gets the nesting level of the table. Top-level tables have level 1.
- parentBody — For EACH ITEM in the collection: Gets the parent body of the table.
- parentContentControl — For EACH ITEM in the collection: Gets the content control that contains the table. Throws an ItemNotFound error if there isn't a parent content control.
- parentContentControlOrNullObject — For EACH ITEM in the collection: Gets the content control that contains the table. If there isn't a parent content control, then this method will return an object with its isNullObject property set to true. For further information, see OrNullObject methods and properties.
- parentTable — For EACH ITEM in the collection: Gets the table that contains this table. Throws an ItemNotFound error if it isn't contained in a table.
- parentTableCell — For EACH ITEM in the collection: Gets the table cell that contains this table. Throws an ItemNotFound error if it isn't contained in a table cell.
- parentTableCellOrNullObject — For EACH ITEM in the collection: Gets the table cell that contains this table. If it isn't contained in a table cell, then this method will return an object with its isNullObject property set to true. For further information, see OrNullObject methods and properties.
- parentTableOrNullObject — For EACH ITEM in the collection: Gets the table that contains this table. If it isn't contained in a table, then this method will return an object with its isNullObject property set to true. For further information, see OrNullObject methods and properties.
- rowCount — For EACH ITEM in the collection: Gets the number of rows in the table.
- shadingColor — For EACH ITEM in the collection: Specifies the shading color. Color is specified in "#RRGGBB" format or by using the color name.
- style — For EACH ITEM in the collection: Specifies the style name for the table. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
- styleBandedColumns — For EACH ITEM in the collection: Specifies whether the table has banded columns.
- styleBandedRows — For EACH ITEM in the collection: Specifies whether the table has banded rows.
- styleBuiltIn — For EACH ITEM in the collection: Specifies the built-in style name for the table. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
- styleFirstColumn — For EACH ITEM in the collection: Specifies whether the table has a first column with a special style.
- styleLastColumn — For EACH ITEM in the collection: Specifies whether the table has a last column with a special style.
- styleTotalRow — For EACH ITEM in the collection: Specifies whether the table has a total (last) row with a special style.
- values — For EACH ITEM in the collection: Specifies the text values in the table, as a 2D JavaScript array.
- verticalAlignment — For EACH ITEM in the collection: Specifies the vertical alignment of every cell in the table. The value can be 'Top', 'Center', or 'Bottom'.
- width — For EACH ITEM in the collection: Specifies the width of the table in points.

## Property Details

### $all
Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).

```typescript
$all?: boolean;
```

Property value: boolean

---

### alignment
For EACH ITEM in the collection: Specifies the alignment of the table against the page column. The value can be 'Left', 'Centered', or 'Right'.

```typescript
alignment?: boolean;
```

Property value: boolean

Remarks: [API set: WordApi 1.3]

---

### font
For EACH ITEM in the collection: Gets the font. Use this to get and set font name, size, color, and other properties.

```typescript
font?: Word.Interfaces.FontLoadOptions;
```

Property value: [Word.Interfaces.FontLoadOptions](/en-us/javascript/api/word/word.interfaces.fontloadoptions)

Remarks: [API set: WordApi 1.3]

---

### headerRowCount
For EACH ITEM in the collection: Specifies the number of header rows.

```typescript
headerRowCount?: boolean;
```

Property value: boolean

Remarks: [API set: WordApi 1.3]

---

### horizontalAlignment
For EACH ITEM in the collection: Specifies the horizontal alignment of every cell in the table. The value can be 'Left', 'Centered', 'Right', or 'Justified'.

```typescript
horizontalAlignment?: boolean;
```

Property value: boolean

Remarks: [API set: WordApi 1.3]

---

### isUniform
For EACH ITEM in the collection: Indicates whether all of the table rows are uniform.

```typescript
isUniform?: boolean;
```

Property value: boolean

Remarks: [API set: WordApi 1.3]

---

### nestingLevel
For EACH ITEM in the collection: Gets the nesting level of the table. Top-level tables have level 1.

```typescript
nestingLevel?: boolean;
```

Property value: boolean

Remarks: [API set: WordApi 1.3]

---

### parentBody
For EACH ITEM in the collection: Gets the parent body of the table.

```typescript
parentBody?: Word.Interfaces.BodyLoadOptions;
```

Property value: [Word.Interfaces.BodyLoadOptions](/en-us/javascript/api/word/word.interfaces.bodyloadoptions)

Remarks: [API set: WordApi 1.3]

---

### parentContentControl
For EACH ITEM in the collection: Gets the content control that contains the table. Throws an ItemNotFound error if there isn't a parent content control.

```typescript
parentContentControl?: Word.Interfaces.ContentControlLoadOptions;
```

Property value: [Word.Interfaces.ContentControlLoadOptions](/en-us/javascript/api/word/word.interfaces.contentcontrolloadoptions)

Remarks: [API set: WordApi 1.3]

---

### parentContentControlOrNullObject
For EACH ITEM in the collection: Gets the content control that contains the table. If there isn't a parent content control, then this method will return an object with its isNullObject property set to true. For further information, see OrNullObject methods and properties.

```typescript
parentContentControlOrNullObject?: Word.Interfaces.ContentControlLoadOptions;
```

Property value: [Word.Interfaces.ContentControlLoadOptions](/en-us/javascript/api/word/word.interfaces.contentcontrolloadoptions)

Remarks: [API set: WordApi 1.3]

---

### parentTable
For EACH ITEM in the collection: Gets the table that contains this table. Throws an ItemNotFound error if it isn't contained in a table.

```typescript
parentTable?: Word.Interfaces.TableLoadOptions;
```

Property value: [Word.Interfaces.TableLoadOptions](/en-us/javascript/api/word/word.interfaces.tableloadoptions)

Remarks: [API set: WordApi 1.3]

---

### parentTableCell
For EACH ITEM in the collection: Gets the table cell that contains this table. Throws an ItemNotFound error if it isn't contained in a table cell.

```typescript
parentTableCell?: Word.Interfaces.TableCellLoadOptions;
```

Property value: [Word.Interfaces.TableCellLoadOptions](/en-us/javascript/api/word/word.interfaces.tablecellloadoptions)

Remarks: [API set: WordApi 1.3]

---

### parentTableCellOrNullObject
For EACH ITEM in the collection: Gets the table cell that contains this table. If it isn't contained in a table cell, then this method will return an object with its isNullObject property set to true. For further information, see OrNullObject methods and properties.

```typescript
parentTableCellOrNullObject?: Word.Interfaces.TableCellLoadOptions;
```

Property value: [Word.Interfaces.TableCellLoadOptions](/en-us/javascript/api/word/word.interfaces.tablecellloadoptions)

Remarks: [API set: WordApi 1.3]

---

### parentTableOrNullObject
For EACH ITEM in the collection: Gets the table that contains this table. If it isn't contained in a table, then this method will return an object with its isNullObject property set to true. For further information, see OrNullObject methods and properties.

```typescript
parentTableOrNullObject?: Word.Interfaces.TableLoadOptions;
```

Property value: [Word.Interfaces.TableLoadOptions](/en-us/javascript/api/word/word.interfaces.tableloadoptions)

Remarks: [API set: WordApi 1.3]

---

### rowCount
For EACH ITEM in the collection: Gets the number of rows in the table.

```typescript
rowCount?: boolean;
```

Property value: boolean

Remarks: [API set: WordApi 1.3]

---

### shadingColor
For EACH ITEM in the collection: Specifies the shading color. Color is specified in "#RRGGBB" format or by using the color name.

```typescript
shadingColor?: boolean;
```

Property value: boolean

Remarks: [API set: WordApi 1.3]

---

### style
For EACH ITEM in the collection: Specifies the style name for the table. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.

```typescript
style?: boolean;
```

Property value: boolean

Remarks: [API set: WordApi 1.3]

---

### styleBandedColumns
For EACH ITEM in the collection: Specifies whether the table has banded columns.

```typescript
styleBandedColumns?: boolean;
```

Property value: boolean

Remarks: [API set: WordApi 1.3]

---

### styleBandedRows
For EACH ITEM in the collection: Specifies whether the table has banded rows.

```typescript
styleBandedRows?: boolean;
```

Property value: boolean

Remarks: [API set: WordApi 1.3]

---

### styleBuiltIn
For EACH ITEM in the collection: Specifies the built-in style name for the table. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.

```typescript
styleBuiltIn?: boolean;
```

Property value: boolean

Remarks: [API set: WordApi 1.3]

---

### styleFirstColumn
For EACH ITEM in the collection: Specifies whether the table has a first column with a special style.

```typescript
styleFirstColumn?: boolean;
```

Property value: boolean

Remarks: [API set: WordApi 1.3]

---

### styleLastColumn
For EACH ITEM in the collection: Specifies whether the table has a last column with a special style.

```typescript
styleLastColumn?: boolean;
```

Property value: boolean

Remarks: [API set: WordApi 1.3]

---

### styleTotalRow
For EACH ITEM in the collection: Specifies whether the table has a total (last) row with a special style.

```typescript
styleTotalRow?: boolean;
```

Property value: boolean

Remarks: [API set: WordApi 1.3]

---

### values
For EACH ITEM in the collection: Specifies the text values in the table, as a 2D JavaScript array.

```typescript
values?: boolean;
```

Property value: boolean

Remarks: [API set: WordApi 1.3]

---

### verticalAlignment
For EACH ITEM in the collection: Specifies the vertical alignment of every cell in the table. The value can be 'Top', 'Center', or 'Bottom'.

```typescript
verticalAlignment?: boolean;
```

Property value: boolean

Remarks: [API set: WordApi 1.3]

---

### width
For EACH ITEM in the collection: Specifies the width of the table in points.

```typescript
width?: boolean;
```

Property value: boolean

Remarks: [API set: WordApi 1.3]