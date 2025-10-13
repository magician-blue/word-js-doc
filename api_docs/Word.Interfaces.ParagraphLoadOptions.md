# Word.Interfaces.ParagraphLoadOptions interface

- Package: [word](/en-us/javascript/api/word)

Represents a single paragraph in a selection, range, content control, or document body.

## Remarks

[ API set: WordApi 1.1 ]

## Properties

- $all: Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
- alignment: Specifies the alignment for a paragraph. The value can be 'left', 'centered', 'right', or 'justified'.
- firstLineIndent: Specifies the value, in points, for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent.
- font: Gets the text format of the paragraph. Use this to get and set font name, size, color, and other properties.
- isLastParagraph: Indicates the paragraph is the last one inside its parent body.
- isListItem: Checks whether the paragraph is a list item.
- leftIndent: Specifies the left indent value, in points, for the paragraph.
- lineSpacing: Specifies the line spacing, in points, for the specified paragraph. In the Word UI, this value is divided by 12.
- lineUnitAfter: Specifies the amount of spacing, in grid lines, after the paragraph.
- lineUnitBefore: Specifies the amount of spacing, in grid lines, before the paragraph.
- list: Gets the List to which this paragraph belongs. Throws an `ItemNotFound` error if the paragraph isn't in a list.
- listItem: Gets the ListItem for the paragraph. Throws an `ItemNotFound` error if the paragraph isn't part of a list.
- listItemOrNullObject: Gets the ListItem for the paragraph. If the paragraph isn't part of a list, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).
- listOrNullObject: Gets the List to which this paragraph belongs. If the paragraph isn't in a list, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).
- outlineLevel: Specifies the outline level for the paragraph.
- parentBody: Gets the parent body of the paragraph.
- parentContentControl: Gets the content control that contains the paragraph. Throws an `ItemNotFound` error if there isn't a parent content control.
- parentContentControlOrNullObject: Gets the content control that contains the paragraph. If there isn't a parent content control, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).
- parentTable: Gets the table that contains the paragraph. Throws an `ItemNotFound` error if it isn't contained in a table.
- parentTableCell: Gets the table cell that contains the paragraph. Throws an `ItemNotFound` error if it isn't contained in a table cell.
- parentTableCellOrNullObject: Gets the table cell that contains the paragraph. If it isn't contained in a table cell, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).
- parentTableOrNullObject: Gets the table that contains the paragraph. If it isn't contained in a table, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).
- rightIndent: Specifies the right indent value, in points, for the paragraph.
- shading: Returns a `ShadingUniversal` object that refers to the shading formatting for the paragraph.
- spaceAfter: Specifies the spacing, in points, after the paragraph.
- spaceBefore: Specifies the spacing, in points, before the paragraph.
- style: Specifies the style name for the paragraph. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
- styleBuiltIn: Specifies the built-in style name for the paragraph. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
- tableNestingLevel: Gets the level of the paragraph's table. It returns 0 if the paragraph isn't in a table.
- text: Gets the text of the paragraph.
- uniqueLocalId: Gets a string that represents the paragraph identifier in the current session. ID is in standard 8-4-4-4-12 GUID format without curly braces and differs across sessions and coauthors.

## Property Details

### $all

Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

#### Property value

- boolean

---

### alignment

Specifies the alignment for a paragraph. The value can be 'left', 'centered', 'right', or 'justified'.

```typescript
alignment?: boolean;
```

#### Property value

- boolean

#### Remarks

[ API set: WordApi 1.1 ]

---

### firstLineIndent

Specifies the value, in points, for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent.

```typescript
firstLineIndent?: boolean;
```

#### Property value

- boolean

#### Remarks

[ API set: WordApi 1.1 ]

---

### font

Gets the text format of the paragraph. Use this to get and set font name, size, color, and other properties.

```typescript
font?: Word.Interfaces.FontLoadOptions;
```

#### Property value

- [Word.Interfaces.FontLoadOptions](/en-us/javascript/api/word/word.interfaces.fontloadoptions)

#### Remarks

[ API set: WordApi 1.1 ]

---

### isLastParagraph

Indicates the paragraph is the last one inside its parent body.

```typescript
isLastParagraph?: boolean;
```

#### Property value

- boolean

#### Remarks

[ API set: WordApi 1.3 ]

---

### isListItem

Checks whether the paragraph is a list item.

```typescript
isListItem?: boolean;
```

#### Property value

- boolean

#### Remarks

[ API set: WordApi 1.3 ]

---

### leftIndent

Specifies the left indent value, in points, for the paragraph.

```typescript
leftIndent?: boolean;
```

#### Property value

- boolean

#### Remarks

[ API set: WordApi 1.1 ]

---

### lineSpacing

Specifies the line spacing, in points, for the specified paragraph. In the Word UI, this value is divided by 12.

```typescript
lineSpacing?: boolean;
```

#### Property value

- boolean

#### Remarks

[ API set: WordApi 1.1 ]

---

### lineUnitAfter

Specifies the amount of spacing, in grid lines, after the paragraph.

```typescript
lineUnitAfter?: boolean;
```

#### Property value

- boolean

#### Remarks

[ API set: WordApi 1.1 ]

---

### lineUnitBefore

Specifies the amount of spacing, in grid lines, before the paragraph.

```typescript
lineUnitBefore?: boolean;
```

#### Property value

- boolean

#### Remarks

[ API set: WordApi 1.1 ]

---

### list

Gets the List to which this paragraph belongs. Throws an `ItemNotFound` error if the paragraph isn't in a list.

```typescript
list?: Word.Interfaces.ListLoadOptions;
```

#### Property value

- [Word.Interfaces.ListLoadOptions](/en-us/javascript/api/word/word.interfaces.listloadoptions)

#### Remarks

[ API set: WordApi 1.3 ]

---

### listItem

Gets the ListItem for the paragraph. Throws an `ItemNotFound` error if the paragraph isn't part of a list.

```typescript
listItem?: Word.Interfaces.ListItemLoadOptions;
```

#### Property value

- [Word.Interfaces.ListItemLoadOptions](/en-us/javascript/api/word/word.interfaces.listitemloadoptions)

#### Remarks

[ API set: WordApi 1.3 ]

---

### listItemOrNullObject

Gets the ListItem for the paragraph. If the paragraph isn't part of a list, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
listItemOrNullObject?: Word.Interfaces.ListItemLoadOptions;
```

#### Property value

- [Word.Interfaces.ListItemLoadOptions](/en-us/javascript/api/word/word.interfaces.listitemloadoptions)

#### Remarks

[ API set: WordApi 1.3 ]

---

### listOrNullObject

Gets the List to which this paragraph belongs. If the paragraph isn't in a list, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
listOrNullObject?: Word.Interfaces.ListLoadOptions;
```

#### Property value

- [Word.Interfaces.ListLoadOptions](/en-us/javascript/api/word/word.interfaces.listloadoptions)

#### Remarks

[ API set: WordApi 1.3 ]

---

### outlineLevel

Specifies the outline level for the paragraph.

```typescript
outlineLevel?: boolean;
```

#### Property value

- boolean

#### Remarks

[ API set: WordApi 1.1 ]

---

### parentBody

Gets the parent body of the paragraph.

```typescript
parentBody?: Word.Interfaces.BodyLoadOptions;
```

#### Property value

- [Word.Interfaces.BodyLoadOptions](/en-us/javascript/api/word/word.interfaces.bodyloadoptions)

#### Remarks

[ API set: WordApi 1.3 ]

---

### parentContentControl

Gets the content control that contains the paragraph. Throws an `ItemNotFound` error if there isn't a parent content control.

```typescript
parentContentControl?: Word.Interfaces.ContentControlLoadOptions;
```

#### Property value

- [Word.Interfaces.ContentControlLoadOptions](/en-us/javascript/api/word/word.interfaces.contentcontrolloadoptions)

#### Remarks

[ API set: WordApi 1.1 ]

---

### parentContentControlOrNullObject

Gets the content control that contains the paragraph. If there isn't a parent content control, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
parentContentControlOrNullObject?: Word.Interfaces.ContentControlLoadOptions;
```

#### Property value

- [Word.Interfaces.ContentControlLoadOptions](/en-us/javascript/api/word/word.interfaces.contentcontrolloadoptions)

#### Remarks

[ API set: WordApi 1.3 ]

---

### parentTable

Gets the table that contains the paragraph. Throws an `ItemNotFound` error if it isn't contained in a table.

```typescript
parentTable?: Word.Interfaces.TableLoadOptions;
```

#### Property value

- [Word.Interfaces.TableLoadOptions](/en-us/javascript/api/word/word.interfaces.tableloadoptions)

#### Remarks

[ API set: WordApi 1.3 ]

---

### parentTableCell

Gets the table cell that contains the paragraph. Throws an `ItemNotFound` error if it isn't contained in a table cell.

```typescript
parentTableCell?: Word.Interfaces.TableCellLoadOptions;
```

#### Property value

- [Word.Interfaces.TableCellLoadOptions](/en-us/javascript/api/word/word.interfaces.tablecellloadoptions)

#### Remarks

[ API set: WordApi 1.3 ]

---

### parentTableCellOrNullObject

Gets the table cell that contains the paragraph. If it isn't contained in a table cell, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
parentTableCellOrNullObject?: Word.Interfaces.TableCellLoadOptions;
```

#### Property value

- [Word.Interfaces.TableCellLoadOptions](/en-us/javascript/api/word/word.interfaces.tablecellloadoptions)

#### Remarks

[ API set: WordApi 1.3 ]

---

### parentTableOrNullObject

Gets the table that contains the paragraph. If it isn't contained in a table, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
parentTableOrNullObject?: Word.Interfaces.TableLoadOptions;
```

#### Property value

- [Word.Interfaces.TableLoadOptions](/en-us/javascript/api/word/word.interfaces.tableloadoptions)

#### Remarks

[ API set: WordApi 1.3 ]

---

### rightIndent

Specifies the right indent value, in points, for the paragraph.

```typescript
rightIndent?: boolean;
```

#### Property value

- boolean

#### Remarks

[ API set: WordApi 1.1 ]

---

### shading

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `ShadingUniversal` object that refers to the shading formatting for the paragraph.

```typescript
shading?: Word.Interfaces.ShadingUniversalLoadOptions;
```

#### Property value

- [Word.Interfaces.ShadingUniversalLoadOptions](/en-us/javascript/api/word/word.interfaces.shadinguniversalloadoptions)

#### Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### spaceAfter

Specifies the spacing, in points, after the paragraph.

```typescript
spaceAfter?: boolean;
```

#### Property value

- boolean

#### Remarks

[ API set: WordApi 1.1 ]

---

### spaceBefore

Specifies the spacing, in points, before the paragraph.

```typescript
spaceBefore?: boolean;
```

#### Property value

- boolean

#### Remarks

[ API set: WordApi 1.1 ]

---

### style

Specifies the style name for the paragraph. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.

```typescript
style?: boolean;
```

#### Property value

- boolean

#### Remarks

[ API set: WordApi 1.1 ]

---

### styleBuiltIn

Specifies the built-in style name for the paragraph. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.

```typescript
styleBuiltIn?: boolean;
```

#### Property value

- boolean

#### Remarks

[ API set: WordApi 1.3 ]

---

### tableNestingLevel

Gets the level of the paragraph's table. It returns 0 if the paragraph isn't in a table.

```typescript
tableNestingLevel?: boolean;
```

#### Property value

- boolean

#### Remarks

[ API set: WordApi 1.3 ]

---

### text

Gets the text of the paragraph.

```typescript
text?: boolean;
```

#### Property value

- boolean

#### Remarks

[ API set: WordApi 1.1 ]

---

### uniqueLocalId

Gets a string that represents the paragraph identifier in the current session. ID is in standard 8-4-4-4-12 GUID format without curly braces and differs across sessions and coauthors.

```typescript
uniqueLocalId?: boolean;
```

#### Property value

- boolean

#### Remarks

[ API set: WordApi 1.6 ]