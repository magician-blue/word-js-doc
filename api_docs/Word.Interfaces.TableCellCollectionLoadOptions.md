# Word.Interfaces.TableCellCollectionLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Contains the collection of the document's TableCell objects.

## Remarks

[API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- $all  
  Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

- body  
  For EACH ITEM in the collection: Gets the body object of the cell.

- cellIndex  
  For EACH ITEM in the collection: Gets the index of the cell in its row.

- columnWidth  
  For EACH ITEM in the collection: Specifies the width of the cell's column in points. This is applicable to uniform tables.

- horizontalAlignment  
  For EACH ITEM in the collection: Specifies the horizontal alignment of the cell. The value can be 'Left', 'Centered', 'Right', or 'Justified'.

- parentRow  
  For EACH ITEM in the collection: Gets the parent row of the cell.

- parentTable  
  For EACH ITEM in the collection: Gets the parent table of the cell.

- rowIndex  
  For EACH ITEM in the collection: Gets the index of the cell's row in the table.

- shadingColor  
  For EACH ITEM in the collection: Specifies the shading color of the cell. Color is specified in "#RRGGBB" format or by using the color name.

- value  
  For EACH ITEM in the collection: Specifies the text of the cell.

- verticalAlignment  
  For EACH ITEM in the collection: Specifies the vertical alignment of the cell. The value can be 'Top', 'Center', or 'Bottom'.

- width  
  For EACH ITEM in the collection: Gets the width of the cell in points.

## Property Details

### $all

Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

Property value: boolean

---

### body

For EACH ITEM in the collection: Gets the body object of the cell.

```typescript
body?: Word.Interfaces.BodyLoadOptions;
```

Property value: [Word.Interfaces.BodyLoadOptions](/en-us/javascript/api/word/word.interfaces.bodyloadoptions)

Remarks  
[API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### cellIndex

For EACH ITEM in the collection: Gets the index of the cell in its row.

```typescript
cellIndex?: boolean;
```

Property value: boolean

Remarks  
[API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### columnWidth

For EACH ITEM in the collection: Specifies the width of the cell's column in points. This is applicable to uniform tables.

```typescript
columnWidth?: boolean;
```

Property value: boolean

Remarks  
[API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### horizontalAlignment

For EACH ITEM in the collection: Specifies the horizontal alignment of the cell. The value can be 'Left', 'Centered', 'Right', or 'Justified'.

```typescript
horizontalAlignment?: boolean;
```

Property value: boolean

Remarks  
[API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### parentRow

For EACH ITEM in the collection: Gets the parent row of the cell.

```typescript
parentRow?: Word.Interfaces.TableRowLoadOptions;
```

Property value: [Word.Interfaces.TableRowLoadOptions](/en-us/javascript/api/word/word.interfaces.tablerowloadoptions)

Remarks  
[API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### parentTable

For EACH ITEM in the collection: Gets the parent table of the cell.

```typescript
parentTable?: Word.Interfaces.TableLoadOptions;
```

Property value: [Word.Interfaces.TableLoadOptions](/en-us/javascript/api/word/word.interfaces.tableloadoptions)

Remarks  
[API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### rowIndex

For EACH ITEM in the collection: Gets the index of the cell's row in the table.

```typescript
rowIndex?: boolean;
```

Property value: boolean

Remarks  
[API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### shadingColor

For EACH ITEM in the collection: Specifies the shading color of the cell. Color is specified in "#RRGGBB" format or by using the color name.

```typescript
shadingColor?: boolean;
```

Property value: boolean

Remarks  
[API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### value

For EACH ITEM in the collection: Specifies the text of the cell.

```typescript
value?: boolean;
```

Property value: boolean

Remarks  
[API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### verticalAlignment

For EACH ITEM in the collection: Specifies the vertical alignment of the cell. The value can be 'Top', 'Center', or 'Bottom'.

```typescript
verticalAlignment?: boolean;
```

Property value: boolean

Remarks  
[API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### width

For EACH ITEM in the collection: Gets the width of the cell in points.

```typescript
width?: boolean;
```

Property value: boolean

Remarks  
[API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)