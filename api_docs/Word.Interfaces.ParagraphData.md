# Word.Interfaces.ParagraphData interface

Package: https://learn.microsoft.com/en-us/javascript/api/word

An interface describing the data returned by calling paragraph.toJSON().

## Properties

- alignment — Specifies the alignment for a paragraph. The value can be 'left', 'centered', 'right', or 'justified'.
- borders — Returns a BorderUniversalCollection object that represents all the borders for the paragraph.
- fields — Gets the collection of fields in the paragraph.
- firstLineIndent — Specifies the value, in points, for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent.
- font — Gets the text format of the paragraph. Use this to get and set font name, size, color, and other properties.
- inlinePictures — Gets the collection of InlinePicture objects in the paragraph. The collection doesn't include floating images.
- isLastParagraph — Indicates the paragraph is the last one inside its parent body.
- isListItem — Checks whether the paragraph is a list item.
- leftIndent — Specifies the left indent value, in points, for the paragraph.
- lineSpacing — Specifies the line spacing, in points, for the specified paragraph. In the Word UI, this value is divided by 12.
- lineUnitAfter — Specifies the amount of spacing, in grid lines, after the paragraph.
- lineUnitBefore — Specifies the amount of spacing, in grid lines, before the paragraph.
- listItem — Gets the ListItem for the paragraph. Throws an ItemNotFound error if the paragraph isn't part of a list.
- listItemOrNullObject — Gets the ListItem for the paragraph. If the paragraph isn't part of a list, then this method will return an object with its isNullObject property set to true. For further information, see https://learn.microsoft.com/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties.
- outlineLevel — Specifies the outline level for the paragraph.
- rightIndent — Specifies the right indent value, in points, for the paragraph.
- shading — Returns a ShadingUniversal object that refers to the shading formatting for the paragraph.
- shapes — Gets the collection of shape objects anchored in the paragraph, including both inline and floating shapes. Currently, only the following shapes are supported: text boxes, geometric shapes, groups, pictures, and canvases.
- spaceAfter — Specifies the spacing, in points, after the paragraph.
- spaceBefore — Specifies the spacing, in points, before the paragraph.
- style — Specifies the style name for the paragraph. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
- styleBuiltIn — Specifies the built-in style name for the paragraph. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
- tableNestingLevel — Gets the level of the paragraph's table. It returns 0 if the paragraph isn't in a table.
- text — Gets the text of the paragraph.
- uniqueLocalId — Gets a string that represents the paragraph identifier in the current session. ID is in standard 8-4-4-4-12 GUID format without curly braces and differs across sessions and coauthors.

## Property Details

### alignment

Specifies the alignment for a paragraph. The value can be 'left', 'centered', 'right', or 'justified'.

```typescript
alignment?: Word.Alignment | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified";
```

- Property Value: https://learn.microsoft.com/en-us/javascript/api/word/word.alignment | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified"
- Remarks: [ API set: WordApi 1.1 ]

### borders

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a BorderUniversalCollection object that represents all the borders for the paragraph.

```typescript
borders?: Word.Interfaces.BorderUniversalData[];
```

- Property Value: https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.borderuniversaldata[]
- Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

### fields

Gets the collection of fields in the paragraph.

```typescript
fields?: Word.Interfaces.FieldData[];
```

- Property Value: https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.fielddata[]
- Remarks: [ API set: WordApi 1.4 ]

### firstLineIndent

Specifies the value, in points, for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent.

```typescript
firstLineIndent?: number;
```

- Property Value: number
- Remarks: [ API set: WordApi 1.1 ]

### font

Gets the text format of the paragraph. Use this to get and set font name, size, color, and other properties.

```typescript
font?: Word.Interfaces.FontData;
```

- Property Value: https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.fontdata
- Remarks: [ API set: WordApi 1.1 ]

### inlinePictures

Gets the collection of InlinePicture objects in the paragraph. The collection doesn't include floating images.

```typescript
inlinePictures?: Word.Interfaces.InlinePictureData[];
```

- Property Value: https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.inlinepicturedata[]
- Remarks: [ API set: WordApi 1.1 ]

### isLastParagraph

Indicates the paragraph is the last one inside its parent body.

```typescript
isLastParagraph?: boolean;
```

- Property Value: boolean
- Remarks: [ API set: WordApi 1.3 ]

### isListItem

Checks whether the paragraph is a list item.

```typescript
isListItem?: boolean;
```

- Property Value: boolean
- Remarks: [ API set: WordApi 1.3 ]

### leftIndent

Specifies the left indent value, in points, for the paragraph.

```typescript
leftIndent?: number;
```

- Property Value: number
- Remarks: [ API set: WordApi 1.1 ]

### lineSpacing

Specifies the line spacing, in points, for the specified paragraph. In the Word UI, this value is divided by 12.

```typescript
lineSpacing?: number;
```

- Property Value: number
- Remarks: [ API set: WordApi 1.1 ]

### lineUnitAfter

Specifies the amount of spacing, in grid lines, after the paragraph.

```typescript
lineUnitAfter?: number;
```

- Property Value: number
- Remarks: [ API set: WordApi 1.1 ]

### lineUnitBefore

Specifies the amount of spacing, in grid lines, before the paragraph.

```typescript
lineUnitBefore?: number;
```

- Property Value: number
- Remarks: [ API set: WordApi 1.1 ]

### listItem

Gets the ListItem for the paragraph. Throws an ItemNotFound error if the paragraph isn't part of a list.

```typescript
listItem?: Word.Interfaces.ListItemData;
```

- Property Value: https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.listitemdata
- Remarks: [ API set: WordApi 1.3 ]

### listItemOrNullObject

Gets the ListItem for the paragraph. If the paragraph isn't part of a list, then this method will return an object with its isNullObject property set to true. For further information, see https://learn.microsoft.com/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties.

```typescript
listItemOrNullObject?: Word.Interfaces.ListItemData;
```

- Property Value: https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.listitemdata
- Remarks: [ API set: WordApi 1.3 ]

### outlineLevel

Specifies the outline level for the paragraph.

```typescript
outlineLevel?: number;
```

- Property Value: number
- Remarks: [ API set: WordApi 1.1 ]

### rightIndent

Specifies the right indent value, in points, for the paragraph.

```typescript
rightIndent?: number;
```

- Property Value: number
- Remarks: [ API set: WordApi 1.1 ]

### shading

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a ShadingUniversal object that refers to the shading formatting for the paragraph.

```typescript
shading?: Word.Interfaces.ShadingUniversalData;
```

- Property Value: https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.shadinguniversaldata
- Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

### shapes

Gets the collection of shape objects anchored in the paragraph, including both inline and floating shapes. Currently, only the following shapes are supported: text boxes, geometric shapes, groups, pictures, and canvases.

```typescript
shapes?: Word.Interfaces.ShapeData[];
```

- Property Value: https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.shapedata[]
- Remarks: [ API set: WordApiDesktop 1.2 ]

### spaceAfter

Specifies the spacing, in points, after the paragraph.

```typescript
spaceAfter?: number;
```

- Property Value: number
- Remarks: [ API set: WordApi 1.1 ]

### spaceBefore

Specifies the spacing, in points, before the paragraph.

```typescript
spaceBefore?: number;
```

- Property Value: number
- Remarks: [ API set: WordApi 1.1 ]

### style

Specifies the style name for the paragraph. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.

```typescript
style?: string;
```

- Property Value: string
- Remarks: [ API set: WordApi 1.1 ]

### styleBuiltIn

Specifies the built-in style name for the paragraph. Use this property for built-in styles that are