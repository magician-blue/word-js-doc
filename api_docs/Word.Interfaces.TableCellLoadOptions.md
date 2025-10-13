# Word.Interfaces.TableCellLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Represents a table cell in a Word document.

## Remarks

[API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- [$all](#word-word-interfaces-tablecellloadoptions-all-member): Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
- [body](#word-word-interfaces-tablecellloadoptions-body-member): Gets the body object of the cell.
- [cellIndex](#word-word-interfaces-tablecellloadoptions-cellindex-member): Gets the index of the cell in its row.
- [columnWidth](#word-word-interfaces-tablecellloadoptions-columnwidth-member): Specifies the width of the cell's column in points. This is applicable to uniform tables.
- [horizontalAlignment](#word-word-interfaces-tablecellloadoptions-horizontalalignment-member): Specifies the horizontal alignment of the cell. The value can be 'Left', 'Centered', 'Right', or 'Justified'.
- [parentRow](#word-word-interfaces-tablecellloadoptions-parentrow-member): Gets the parent row of the cell.
- [parentTable](#word-word-interfaces-tablecellloadoptions-parenttable-member): Gets the parent table of the cell.
- [rowIndex](#word-word-interfaces-tablecellloadoptions-rowindex-member): Gets the index of the cell's row in the table.
- [shadingColor](#word-word-interfaces-tablecellloadoptions-shadingcolor-member): Specifies the shading color of the cell. Color is specified in "#RRGGBB" format or by using the color name.
- [value](#word-word-interfaces-tablecellloadoptions-value-member): Specifies the text of the cell.
- [verticalAlignment](#word-word-interfaces-tablecellloadoptions-verticalalignment-member): Specifies the vertical alignment of the cell. The value can be 'Top', 'Center', or 'Bottom'.
- [width](#word-word-interfaces-tablecellloadoptions-width-member): Gets the width of the cell in points.

## Property Details

<a id="word-word-interfaces-tablecellloadoptions-all-member"></a>
### $all

Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

#### Property Value
- boolean

<a id="word-word-interfaces-tablecellloadoptions-body-member"></a>
### body

Gets the body object of the cell.

```typescript
body?: Word.Interfaces.BodyLoadOptions;
```

#### Property Value
- [Word.Interfaces.BodyLoadOptions](/en-us/javascript/api/word/word.interfaces.bodyloadoptions)

#### Remarks
[API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-tablecellloadoptions-cellindex-member"></a>
### cellIndex

Gets the index of the cell in its row.

```typescript
cellIndex?: boolean;
```

#### Property Value
- boolean

#### Remarks
[API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-tablecellloadoptions-columnwidth-member"></a>
### columnWidth

Specifies the width of the cell's column in points. This is applicable to uniform tables.

```typescript
columnWidth?: boolean;
```

#### Property Value
- boolean

#### Remarks
[API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-tablecellloadoptions-horizontalalignment-member"></a>
### horizontalAlignment

Specifies the horizontal alignment of the cell. The value can be 'Left', 'Centered', 'Right', or 'Justified'.

```typescript
horizontalAlignment?: boolean;
```

#### Property Value
- boolean

#### Remarks
[API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-tablecellloadoptions-parentrow-member"></a>
### parentRow

Gets the parent row of the cell.

```typescript
parentRow?: Word.Interfaces.TableRowLoadOptions;
```

#### Property Value
- [Word.Interfaces.TableRowLoadOptions](/en-us/javascript/api/word/word.interfaces.tablerowloadoptions)

#### Remarks
[API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-tablecellloadoptions-parenttable-member"></a>
### parentTable

Gets the parent table of the cell.

```typescript
parentTable?: Word.Interfaces.TableLoadOptions;
```

#### Property Value
- [Word.Interfaces.TableLoadOptions](/en-us/javascript/api/word/word.interfaces.tableloadoptions)

#### Remarks
[API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-tablecellloadoptions-rowindex-member"></a>
### rowIndex

Gets the index of the cell's row in the table.

```typescript
rowIndex?: boolean;
```

#### Property Value
- boolean

#### Remarks
[API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-tablecellloadoptions-shadingcolor-member"></a>
### shadingColor

Specifies the shading color of the cell. Color is specified in "#RRGGBB" format or by using the color name.

```typescript
shadingColor?: boolean;
```

#### Property Value
- boolean

#### Remarks
[API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-tablecellloadoptions-value-member"></a>
### value

Specifies the text of the cell.

```typescript
value?: boolean;
```

#### Property Value
- boolean

#### Remarks
[API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-tablecellloadoptions-verticalalignment-member"></a>
### verticalAlignment

Specifies the vertical alignment of the cell. The value can be 'Top', 'Center', or 'Bottom'.

```typescript
verticalAlignment?: boolean;
```

#### Property Value
- boolean

#### Remarks
[API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-tablecellloadoptions-width-member"></a>
### width

Gets the width of the cell in points.

```typescript
width?: boolean;
```

#### Property Value
- boolean

#### Remarks
[API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)