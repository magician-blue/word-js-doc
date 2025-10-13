# Word.Interfaces.TableLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Represents a table in a Word document.

## Remarks
[ API set: WordApi 1.3 ]

## Properties
- $all: Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
- alignment: Specifies the alignment of the table against the page column. The value can be 'Left', 'Centered', or 'Right'.
- font: Gets the font. Use this to get and set font name, size, color, and other properties.
- headerRowCount: Specifies the number of header rows.
- horizontalAlignment: Specifies the horizontal alignment of every cell in the table. The value can be 'Left', 'Centered', 'Right', or 'Justified'.
- isUniform: Indicates whether all of the table rows are uniform.
- nestingLevel: Gets the nesting level of the table. Top-level tables have level 1.
- parentBody: Gets the parent body of the table.
- parentContentControl: Gets the content control that contains the table. Throws an `ItemNotFound` error if there isn't a parent content control.
- parentContentControlOrNullObject: Gets the content control that contains the table. If there isn't a parent content control, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [\*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).
- parentTable: Gets the table that contains this table. Throws an `ItemNotFound` error if it isn't contained in a table.
- parentTableCell: Gets the table cell that contains this table. Throws an `ItemNotFound` error if it isn't contained in a table cell.
- parentTableCellOrNullObject: Gets the table cell that contains this table. If it isn't contained in a table cell, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [\*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).
- parentTableOrNullObject: Gets the table that contains this table. If it isn't contained in a table, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [\*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).
- rowCount: Gets the number of rows in the table.
- shadingColor: Specifies the shading color. Color is specified in "#RRGGBB" format or by using the color name.
- style: Specifies the style name for the table. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
- styleBandedColumns: Specifies whether the table has banded columns.
- styleBandedRows: Specifies whether the table has banded rows.
- styleBuiltIn: Specifies the built-in style name for the table. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
- styleFirstColumn: Specifies whether the table has a first column with a special style.
- styleLastColumn: Specifies whether the table has a last column with a special style.
- styleTotalRow: Specifies whether the table has a total (last) row with a special style.
- values: Specifies the text values in the table, as a 2D JavaScript array.
- verticalAlignment: Specifies the vertical alignment of every cell in the table. The value can be 'Top', 'Center', or 'Bottom'.
- width: Specifies the width of the table in points.

## Property Details

### $all
Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

Property Value: boolean

---

### alignment
Specifies the alignment of the table against the page column. The value can be 'Left', 'Centered', or 'Right'.

```typescript
alignment?: boolean;
```

Property Value: boolean

Remarks
[ API set: WordApi 1.3 ]

---

### font
Gets the font. Use this to get and set font name, size, color, and other properties.

```typescript
font?: Word.Interfaces.FontLoadOptions;
```

Property Value: [Word.Interfaces.FontLoadOptions](/en-us/javascript/api/word/word.interfaces.fontloadoptions)

Remarks
[ API set: WordApi 1.3 ]

---

### headerRowCount
Specifies the number of header rows.

```typescript
headerRowCount?: boolean;
```

Property Value: boolean

Remarks
[ API set: WordApi 1.3 ]

---

### horizontalAlignment
Specifies the horizontal alignment of every cell in the table. The value can be 'Left', 'Centered', 'Right', or 'Justified'.

```typescript
horizontalAlignment?: boolean;
```

Property Value: boolean

Remarks
[ API set: WordApi 1.3 ]

---

### isUniform
Indicates whether all of the table rows are uniform.

```typescript
isUniform?: boolean;
```

Property Value: boolean

Remarks
[ API set: WordApi 1.3 ]

---

### nestingLevel
Gets the nesting level of the table. Top-level tables have level 1.

```typescript
nestingLevel?: boolean;
```

Property Value: boolean

Remarks
[ API set: WordApi 1.3 ]

---

### parentBody
Gets the parent body of the table.

```typescript
parentBody?: Word.Interfaces.BodyLoadOptions;
```

Property Value: [Word.Interfaces.BodyLoadOptions](/en-us/javascript/api/word/word.interfaces.bodyloadoptions)

Remarks
[ API set: WordApi 1.3 ]

---

### parentContentControl
Gets the content control that contains the table. Throws an `ItemNotFound` error if there isn't a parent content control.

```typescript
parentContentControl?: Word.Interfaces.ContentControlLoadOptions;
```

Property Value: [Word.Interfaces.ContentControlLoadOptions](/en-us/javascript/api/word/word.interfaces.contentcontrolloadoptions)

Remarks
[ API set: WordApi 1.3 ]

---

### parentContentControlOrNullObject
Gets the content control that contains the table. If there isn't a parent content control, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [\*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
parentContentControlOrNullObject?: Word.Interfaces.ContentControlLoadOptions;
```

Property Value: [Word.Interfaces.ContentControlLoadOptions](/en-us/javascript/api/word/word.interfaces.contentcontrolloadoptions)

Remarks
[ API set: WordApi 1.3 ]

---

### parentTable
Gets the table that contains this table. Throws an `ItemNotFound` error if it isn't contained in a table.

```typescript
parentTable?: Word.Interfaces.TableLoadOptions;
```

Property Value: [Word.Interfaces.TableLoadOptions](/en-us/javascript/api/word/word.interfaces.tableloadoptions)

Remarks
[ API set: WordApi 1.3 ]

---

### parentTableCell
Gets the table cell that contains this table. Throws an `ItemNotFound` error if it isn't contained in a table cell.

```typescript
parentTableCell?: Word.Interfaces.TableCellLoadOptions;
```

Property Value: [Word.Interfaces.TableCellLoadOptions](/en-us/javascript/api/word/word.interfaces.tablecellloadoptions)

Remarks
[ API set: WordApi 1.3 ]

---

### parentTableCellOrNullObject
Gets the table cell that contains this table. If it isn't contained in a table cell, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [\*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
parentTableCellOrNullObject?: Word.Interfaces.TableCellLoadOptions;
```

Property Value: [Word.Interfaces.TableCellLoadOptions](/en-us/javascript/api/word/word.interfaces.tablecellloadoptions)

Remarks
[ API set: WordApi 1.3 ]

---

### parentTableOrNullObject
Gets the table that contains this table. If it isn't contained in a table, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [\*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
parentTableOrNullObject?: Word.Interfaces.TableLoadOptions;
```

Property Value: [Word.Interfaces.TableLoadOptions](/en-us/javascript/api/word/word.interfaces.tableloadoptions)

Remarks
[ API set: WordApi 1.3 ]

---

### rowCount
Gets the number of rows in the table.

```typescript
rowCount?: boolean;
```

Property Value: boolean

Remarks
[ API set: WordApi 1.3 ]

---

### shadingColor
Specifies the shading color. Color is specified in "#RRGGBB" format or by using the color name.

```typescript
shadingColor?: boolean;
```

Property Value: boolean

Remarks
[ API set: WordApi 1.3 ]

---

### style
Specifies the style name for the table. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.

```typescript
style?: boolean;
```

Property Value: boolean

Remarks
[ API set: WordApi 1.3 ]

---

### styleBandedColumns
Specifies whether the table has banded columns.

```typescript
styleBandedColumns?: boolean;
```

Property Value: boolean

Remarks
[ API set: WordApi 1.3 ]

---

### styleBandedRows
Specifies whether the table has banded rows.

```typescript
styleBandedRows?: boolean;
```

Property Value: boolean

Remarks
[ API set: WordApi 1.3 ]

---

### styleBuiltIn
Specifies the built-in style name for the table. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.

```typescript
styleBuiltIn?: boolean;
```

Property Value: boolean

Remarks
[ API set: WordApi 1.3 ]

---

### styleFirstColumn
Specifies whether the table has a first column with a special style.

```typescript
styleFirstColumn?: boolean;
```

Property Value: boolean

Remarks
[ API set: WordApi 1.3 ]

---

### styleLastColumn
Specifies whether the table has a last column with a special style.

```typescript
styleLastColumn?: boolean;
```

Property Value: boolean

Remarks
[ API set: WordApi 1.3 ]

---

### styleTotalRow
Specifies whether the table has a total (last) row with a special style.

```typescript
styleTotalRow?: boolean;
```

Property Value: boolean

Remarks
[ API set: WordApi 1.3 ]

---

### values
Specifies the text values in the table, as a 2D JavaScript array.

```typescript
values?: boolean;
```

Property Value: boolean

Remarks
[ API set: WordApi 1.3 ]

---

### verticalAlignment
Specifies the vertical alignment of every cell in the table. The value can be 'Top', 'Center', or 'Bottom'.

```typescript
verticalAlignment?: boolean;
```

Property Value: boolean

Remarks
[ API set: WordApi 1.3 ]

---

### width
Specifies the width of the table in points.

```typescript
width?: boolean;
```

Property Value: boolean

Remarks
[ API set: WordApi 1.3 ]