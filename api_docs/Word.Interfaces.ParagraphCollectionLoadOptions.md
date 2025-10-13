# Word.Interfaces.ParagraphCollectionLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Contains a collection of [Word.Paragraph](/en-us/javascript/api/word/word.paragraph) objects.

## Remarks

[API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- $all — Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).
- alignment — For EACH ITEM in the collection: Specifies the alignment for a paragraph. The value can be 'left', 'centered', 'right', or 'justified'.
- firstLineIndent — For EACH ITEM in the collection: Specifies the value, in points, for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent.
- font — For EACH ITEM in the collection: Gets the text format of the paragraph. Use this to get and set font name, size, color, and other properties.
- isLastParagraph — For EACH ITEM in the collection: Indicates the paragraph is the last one inside its parent body.
- isListItem — For EACH ITEM in the collection: Checks whether the paragraph is a list item.
- leftIndent — For EACH ITEM in the collection: Specifies the left indent value, in points, for the paragraph.
- lineSpacing — For EACH ITEM in the collection: Specifies the line spacing, in points, for the specified paragraph. In the Word UI, this value is divided by 12.
- lineUnitAfter — For EACH ITEM in the collection: Specifies the amount of spacing, in grid lines, after the paragraph.
- lineUnitBefore — For EACH ITEM in the collection: Specifies the amount of spacing, in grid lines, before the paragraph.
- list — For EACH ITEM in the collection: Gets the List to which this paragraph belongs. Throws an ItemNotFound error if the paragraph isn't in a list.
- listItem — For EACH ITEM in the collection: Gets the ListItem for the paragraph. Throws an ItemNotFound error if the paragraph isn't part of a list.
- listItemOrNullObject — For EACH ITEM in the collection: Gets the ListItem for the paragraph. If the paragraph isn't part of a list, then this method will return an object with its isNullObject property set to true. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).
- listOrNullObject — For EACH ITEM in the collection: Gets the List to which this paragraph belongs. If the paragraph isn't in a list, then this method will return an object with its isNullObject property set to true. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).
- outlineLevel — For EACH ITEM in the collection: Specifies the outline level for the paragraph.
- parentBody — For EACH ITEM in the collection: Gets the parent body of the paragraph.
- parentContentControl — For EACH ITEM in the collection: Gets the content control that contains the paragraph. Throws an ItemNotFound error if there isn't a parent content control.
- parentContentControlOrNullObject — For EACH ITEM in the collection: Gets the content control that contains the paragraph. If there isn't a parent content control, then this method will return an object with its isNullObject property set to true. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).
- parentTable — For EACH ITEM in the collection: Gets the table that contains the paragraph. Throws an ItemNotFound error if it isn't contained in a table.
- parentTableCell — For EACH ITEM in the collection: Gets the table cell that contains the paragraph. Throws an ItemNotFound error if it isn't contained in a table cell.
- parentTableCellOrNullObject — For EACH ITEM in the collection: Gets the table cell that contains the paragraph. If it isn't contained in a table cell, then this method will return an object with its isNullObject property set to true. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).
- parentTableOrNullObject — For EACH ITEM in the collection: Gets the table that contains the paragraph. If it isn't contained in a table, then this method will return an object with its isNullObject property set to true. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).
- rightIndent — For EACH ITEM in the collection: Specifies the right indent value, in points, for the paragraph.
- shading — For EACH ITEM in the collection: Returns a ShadingUniversal object that refers to the shading formatting for the paragraph.
- spaceAfter — For EACH ITEM in the collection: Specifies the spacing, in points, after the paragraph.
- spaceBefore — For EACH ITEM in the collection: Specifies the spacing, in points, before the paragraph.
- style — For EACH ITEM in the collection: Specifies the style name for the paragraph. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
- styleBuiltIn — For EACH ITEM in the collection: Specifies the built-in style name for the paragraph. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
- tableNestingLevel — For EACH ITEM in the collection: Gets the level of the paragraph's table. It returns 0 if the paragraph isn't in a table.
- text — For EACH ITEM in the collection: Gets the text of the paragraph.
- uniqueLocalId — For EACH ITEM in the collection: Gets a string that represents the paragraph identifier in the current session. ID is in standard 8-4-4-4-12 GUID format without curly braces and differs across sessions and coauthors.

## Property Details

### $all

Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).

```typescript
$all?: boolean;
```

Property Value: boolean

---

### alignment

For EACH ITEM in the collection: Specifies the alignment for a paragraph. The value can be 'left', 'centered', 'right', or 'justified'.

```typescript
alignment?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### firstLineIndent

For EACH ITEM in the collection: Specifies the value, in points, for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent.

```typescript
firstLineIndent?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### font

For EACH ITEM in the collection: Gets the text format of the paragraph. Use this to get and set font name, size, color, and other properties.

```typescript
font?: Word.Interfaces.FontLoadOptions;
```

Property Value: [Word.Interfaces.FontLoadOptions](/en-us/javascript/api/word/word.interfaces.fontloadoptions)

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### isLastParagraph

For EACH ITEM in the collection: Indicates the paragraph is the last one inside its parent body.

```typescript
isLastParagraph?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### isListItem

For EACH ITEM in the collection: Checks whether the paragraph is a list item.

```typescript
isListItem?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### leftIndent

For EACH ITEM in the collection: Specifies the left indent value, in points, for the paragraph.

```typescript
leftIndent?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### lineSpacing

For EACH ITEM in the collection: Specifies the line spacing, in points, for the specified paragraph. In the Word UI, this value is divided by 12.

```typescript
lineSpacing?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### lineUnitAfter

For EACH ITEM in the collection: Specifies the amount of spacing, in grid lines, after the paragraph.

```typescript
lineUnitAfter?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### lineUnitBefore

For EACH ITEM in the collection: Specifies the amount of spacing, in grid lines, before the paragraph.

```typescript
lineUnitBefore?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### list

For EACH ITEM in the collection: Gets the List to which this paragraph belongs. Throws an ItemNotFound error if the paragraph isn't in a list.

```typescript
list?: Word.Interfaces.ListLoadOptions;
```

Property Value: [Word.Interfaces.ListLoadOptions](/en-us/javascript/api/word/word.interfaces.listloadoptions)

Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### listItem

For EACH ITEM in the collection: Gets the ListItem for the paragraph. Throws an ItemNotFound error if the paragraph isn't part of a list.

```typescript
listItem?: Word.Interfaces.ListItemLoadOptions;
```

Property Value: [Word.Interfaces.ListItemLoadOptions](/en-us/javascript/api/word/word.interfaces.listitemloadoptions)

Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### listItemOrNullObject

For EACH ITEM in the collection: Gets the ListItem for the paragraph. If the paragraph isn't part of a list, then this method will return an object with its isNullObject property set to true. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
listItemOrNullObject?: Word.Interfaces.ListItemLoadOptions;
```

Property Value: [Word.Interfaces.ListItemLoadOptions](/en-us/javascript/api/word/word.interfaces.listitemloadoptions)

Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### listOrNullObject

For EACH ITEM in the collection: Gets the List to which this paragraph belongs. If the paragraph isn't in a list, then this method will return an object with its isNullObject property set to true. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
listOrNullObject?: Word.Interfaces.ListLoadOptions;
```

Property Value: [Word.Interfaces.ListLoadOptions](/en-us/javascript/api/word/word.interfaces.listloadoptions)

Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### outlineLevel

For EACH ITEM in the collection: Specifies the outline level for the paragraph.

```typescript
outlineLevel?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### parentBody

For EACH ITEM in the collection: Gets the parent body of the paragraph.

```typescript
parentBody?: Word.Interfaces.BodyLoadOptions;
```

Property Value: [Word.Interfaces.BodyLoadOptions](/en-us/javascript/api/word/word.interfaces.bodyloadoptions)

Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### parentContentControl

For EACH ITEM in the collection: Gets the content control that contains the paragraph. Throws an ItemNotFound error if there isn't a parent content control.

```typescript
parentContentControl?: Word.Interfaces.ContentControlLoadOptions;
```

Property Value: [Word.Interfaces.ContentControlLoadOptions](/en-us/javascript/api/word/word.interfaces.contentcontrolloadoptions)

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### parentContentControlOrNullObject

For EACH ITEM in the collection: Gets the content control that contains the paragraph. If there isn't a parent content control, then this method will return an object with its isNullObject property set to true. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
parentContentControlOrNullObject?: Word.Interfaces.ContentControlLoadOptions;
```

Property Value: [Word.Interfaces.ContentControlLoadOptions](/en-us/javascript/api/word/word.interfaces.contentcontrolloadoptions)

Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### parentTable

For EACH ITEM in the collection: Gets the table that contains the paragraph. Throws an ItemNotFound error if it isn't contained in a table.

```typescript
parentTable?: Word.Interfaces.TableLoadOptions;
```

Property Value: [Word.Interfaces.TableLoadOptions](/en-us/javascript/api/word/word.interfaces.tableloadoptions)

Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### parentTableCell

For EACH ITEM in the collection: Gets the table cell that contains the paragraph. Throws an ItemNotFound error if it isn't contained in a table cell.

```typescript
parentTableCell?: Word.Interfaces.TableCellLoadOptions;
```

Property Value: [Word.Interfaces.TableCellLoadOptions](/en-us/javascript/api/word/word.interfaces.tablecellloadoptions)

Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### parentTableCellOrNullObject

For EACH ITEM in the collection: Gets the table cell that contains the paragraph. If it isn't contained in a table cell, then this method will return an object with its isNullObject property set to true. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
parentTableCellOrNullObject?: Word.Interfaces.TableCellLoadOptions;
```

Property Value: [Word.Interfaces.TableCellLoadOptions](/en-us/javascript/api/word/word.interfaces.tablecellloadoptions)

Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### parentTableOrNullObject

For EACH ITEM in the collection: Gets the table that contains the paragraph. If it isn't contained in a table, then this method will return an object with its isNullObject property set to true. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
parentTableOrNullObject?: Word.Interfaces.TableLoadOptions;
```

Property Value: [Word.Interfaces.TableLoadOptions](/en-us/javascript/api/word/word.interfaces.tableloadoptions)

Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### rightIndent

For EACH ITEM in the collection: Specifies the right indent value, in points, for the paragraph.

```typescript
rightIndent?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### shading

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Returns a ShadingUniversal object that refers to the shading formatting for the paragraph.

```typescript
shading?: Word.Interfaces.ShadingUniversalLoadOptions;
```

Property Value: [Word.Interfaces.ShadingUniversalLoadOptions](/en-us/javascript/api/word/word.interfaces.shadinguniversalloadoptions)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### spaceAfter

For EACH ITEM in the collection: Specifies the spacing, in points, after the paragraph.

```typescript
spaceAfter?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### spaceBefore

For EACH ITEM in the collection: Specifies the spacing, in points, before the paragraph.

```typescript
spaceBefore?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### style

For EACH ITEM in the collection: Specifies the style name for the paragraph. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.

```typescript
style?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### styleBuiltIn

For EACH ITEM in the collection: Specifies the built-in style name for the paragraph. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.

```typescript
styleBuiltIn?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### tableNestingLevel

For EACH ITEM in the collection: Gets the level of the paragraph's table. It returns 0 if the paragraph isn't in a table.

```typescript
tableNestingLevel?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### text

For EACH ITEM in the collection: Gets the text of the paragraph.

```typescript
text?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### uniqueLocalId

For EACH ITEM in the collection: Gets a string that represents the paragraph identifier in the current session. ID is in standard 8-4-4-4-12 GUID format without curly braces and differs across sessions and coauthors.

```typescript
uniqueLocalId?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.6](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)