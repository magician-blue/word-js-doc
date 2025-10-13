# Word.Interfaces.TableColumnData interface

- Package: [word](https://learn.microsoft.com/en-us/javascript/api/word)

An interface describing the data returned by calling tableColumn.toJSON().

## Properties

- borders: Returns a BorderUniversalCollection object that represents all the borders for the table column.
- columnIndex: Returns the position of this column in a collection.
- isFirst: Returns true if the column or row is the first one in the table; false otherwise.
- isLast: Returns true if the column or row is the last one in the table; false otherwise.
- nestingLevel: Returns the nesting level of the column.
- preferredWidth: Specifies the preferred width (in points or as a percentage of the window width) for the column. The unit of measurement can be specified by the preferredWidthType property.
- preferredWidthType: Specifies the preferred unit of measurement to use for the width of the table column.
- shading: Returns a ShadingUniversal object that refers to the shading formatting for the column.
- width: Specifies the width of the column, in points.

## Property Details

### borders

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a BorderUniversalCollection object that represents all the borders for the table column.

```typescript
borders?: Word.Interfaces.BorderUniversalData[];
```

#### Property Value

[Word.Interfaces.BorderUniversalData](https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.borderuniversaldata)[]

#### Remarks

[ [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

---

### columnIndex

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the position of this column in a collection.

```typescript
columnIndex?: number;
```

#### Property Value

number

#### Remarks

[ [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

---

### isFirst

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns true if the column or row is the first one in the table; false otherwise.

```typescript
isFirst?: boolean;
```

#### Property Value

boolean

#### Remarks

[ [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

---

### isLast

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns true if the column or row is the last one in the table; false otherwise.

```typescript
isLast?: boolean;
```

#### Property Value

boolean

#### Remarks

[ [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

---

### nestingLevel

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the nesting level of the column.

```typescript
nestingLevel?: number;
```

#### Property Value

number

#### Remarks

[ [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

---

### preferredWidth

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the preferred width (in points or as a percentage of the window width) for the column. The unit of measurement can be specified by the preferredWidthType property.

```typescript
preferredWidth?: number;
```

#### Property Value

number

#### Remarks

[ [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

---

### preferredWidthType

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the preferred unit of measurement to use for the width of the table column.

```typescript
preferredWidthType?: Word.PreferredWidthType | "Auto" | "Percent" | "Points";
```

#### Property Value

[Word.PreferredWidthType](https://learn.microsoft.com/en-us/javascript/api/word/word.preferredwidthtype) | "Auto" | "Percent" | "Points"

#### Remarks

[ [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

---

### shading

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a ShadingUniversal object that refers to the shading formatting for the column.

```typescript
shading?: Word.Interfaces.ShadingUniversalData;
```

#### Property Value

[Word.Interfaces.ShadingUniversalData](https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.shadinguniversaldata)

#### Remarks

[ [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

---

### width

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the width of the column, in points.

```typescript
width?: number;
```

#### Property Value

number

#### Remarks

[ [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]