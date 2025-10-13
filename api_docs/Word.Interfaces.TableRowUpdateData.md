# Word.Interfaces.TableRowUpdateData interface

Package: [word](/en-us/javascript/api/word)

An interface for updating data on the `TableRow` object, for use in `tableRow.set({ ... })`.

## Properties

- font — Gets the font. Use this to get and set font name, size, color, and other properties.
- horizontalAlignment — Specifies the horizontal alignment of every cell in the row. The value can be 'Left', 'Centered', 'Right', or 'Justified'.
- preferredHeight — Specifies the preferred height of the row in points.
- shadingColor — Specifies the shading color. Color is specified in "#RRGGBB" format or by using the color name.
- values — Specifies the text values in the row, as a 2D JavaScript array.
- verticalAlignment — Specifies the vertical alignment of the cells in the row. The value can be 'Top', 'Center', or 'Bottom'.

## Property Details

### font

Gets the font. Use this to get and set font name, size, color, and other properties.

```typescript
font?: Word.Interfaces.FontUpdateData;
```

Property Value
- [Word.Interfaces.FontUpdateData](/en-us/javascript/api/word/word.interfaces.fontupdatedata)

Remarks  
[API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### horizontalAlignment

Specifies the horizontal alignment of every cell in the row. The value can be 'Left', 'Centered', 'Right', or 'Justified'.

```typescript
horizontalAlignment?: Word.Alignment | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified";
```

Property Value
- [Word.Alignment](/en-us/javascript/api/word/word.alignment) | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified"

Remarks  
[API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### preferredHeight

Specifies the preferred height of the row in points.

```typescript
preferredHeight?: number;
```

Property Value
- number

Remarks  
[API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### shadingColor

Specifies the shading color. Color is specified in "#RRGGBB" format or by using the color name.

```typescript
shadingColor?: string;
```

Property Value
- string

Remarks  
[API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### values

Specifies the text values in the row, as a 2D JavaScript array.

```typescript
values?: string[][];
```

Property Value
- string[][]

Remarks  
[API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### verticalAlignment

Specifies the vertical alignment of the cells in the row. The value can be 'Top', 'Center', or 'Bottom'.

```typescript
verticalAlignment?: Word.VerticalAlignment | "Mixed" | "Top" | "Center" | "Bottom";
```

Property Value
- [Word.VerticalAlignment](/en-us/javascript/api/word/word.verticalalignment) | "Mixed" | "Top" | "Center" | "Bottom"

Remarks  
[API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)