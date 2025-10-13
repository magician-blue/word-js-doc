# Word.Interfaces.TableColumnUpdateData interface

Package: [word](/en-us/javascript/api/word)

An interface for updating data on the TableColumn object, for use in tableColumn.set({ ... }).

## Properties

- preferredWidth
  - Specifies the preferred width (in points or as a percentage of the window width) for the column. The unit of measurement can be specified by the preferredWidthType property.
- preferredWidthType
  - Specifies the preferred unit of measurement to use for the width of the table column.
- shading
  - Returns a ShadingUniversal object that refers to the shading formatting for the column.
- width
  - Specifies the width of the column, in points.

## Property Details

### preferredWidth

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the preferred width (in points or as a percentage of the window width) for the column. The unit of measurement can be specified by the preferredWidthType property.

```typescript
preferredWidth?: number;
```

Property value
- number

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### preferredWidthType

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the preferred unit of measurement to use for the width of the table column.

```typescript
preferredWidthType?: Word.PreferredWidthType | "Auto" | "Percent" | "Points";
```

Property value
- [Word.PreferredWidthType](/en-us/javascript/api/word/word.preferredwidthtype) | "Auto" | "Percent" | "Points"

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### shading

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a ShadingUniversal object that refers to the shading formatting for the column.

```typescript
shading?: Word.Interfaces.ShadingUniversalUpdateData;
```

Property value
- [Word.Interfaces.ShadingUniversalUpdateData](/en-us/javascript/api/word/word.interfaces.shadinguniversalupdatedata)

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### width

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the width of the column, in points.

```typescript
width?: number;
```

Property value
- number

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)