# Word.Interfaces.TableColumnCollectionLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents a collection of [Word.TableColumn](/en-us/javascript/api/word/word.tablecolumn) objects in a Word document.

## Remarks

[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- `$all`: Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
- `columnIndex`: For EACH ITEM in the collection: Returns the position of this column in a collection.
- `isFirst`: For EACH ITEM in the collection: Returns `true` if the column or row is the first one in the table; `false` otherwise.
- `isLast`: For EACH ITEM in the collection: Returns `true` if the column or row is the last one in the table; `false` otherwise.
- `nestingLevel`: For EACH ITEM in the collection: Returns the nesting level of the column.
- `preferredWidth`: For EACH ITEM in the collection: Specifies the preferred width (in points or as a percentage of the window width) for the column. The unit of measurement can be specified by the `preferredWidthType` property.
- `preferredWidthType`: For EACH ITEM in the collection: Specifies the preferred unit of measurement to use for the width of the table column.
- `shading`: For EACH ITEM in the collection: Returns a `ShadingUniversal` object that refers to the shading formatting for the column.
- `width`: For EACH ITEM in the collection: Specifies the width of the column, in points.

## Property Details

### $all

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

#### Property Value
boolean

---

### columnIndex

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Returns the position of this column in a collection.

```typescript
columnIndex?: boolean;
```

#### Property Value
boolean

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### isFirst

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Returns `true` if the column or row is the first one in the table; `false` otherwise.

```typescript
isFirst?: boolean;
```

#### Property Value
boolean

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### isLast

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Returns `true` if the column or row is the last one in the table; `false` otherwise.

```typescript
isLast?: boolean;
```

#### Property Value
boolean

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### nestingLevel

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Returns the nesting level of the column.

```typescript
nestingLevel?: boolean;
```

#### Property Value
boolean

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### preferredWidth

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies the preferred width (in points or as a percentage of the window width) for the column. The unit of measurement can be specified by the `preferredWidthType` property.

```typescript
preferredWidth?: boolean;
```

#### Property Value
boolean

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### preferredWidthType

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies the preferred unit of measurement to use for the width of the table column.

```typescript
preferredWidthType?: boolean;
```

#### Property Value
boolean

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### shading

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Returns a `ShadingUniversal` object that refers to the shading formatting for the column.

```typescript
shading?: Word.Interfaces.ShadingUniversalLoadOptions;
```

#### Property Value
[Word.Interfaces.ShadingUniversalLoadOptions](/en-us/javascript/api/word/word.interfaces.shadinguniversalloadoptions)

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### width

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies the width of the column, in points.

```typescript
width?: boolean;
```

#### Property Value
boolean

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)