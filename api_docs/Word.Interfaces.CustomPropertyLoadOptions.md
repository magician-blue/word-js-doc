# Word.Interfaces.CustomPropertyLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Represents a custom property.

## Remarks

[API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- $all — Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
- key — Gets the key of the custom property.
- type — Gets the value type of the custom property. Possible values are: String, Number, Date, Boolean.
- value — Specifies the value of the custom property. Note that even though Word on the web and the docx file format allow these properties to be arbitrarily long, the desktop version of Word will truncate string values to 255 16-bit chars (possibly creating invalid unicode by breaking up a surrogate pair).

## Property Details

### $all

Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

Property Value
- boolean

### key

Gets the key of the custom property.

```typescript
key?: boolean;
```

Property Value
- boolean

Remarks
- [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### type

Gets the value type of the custom property. Possible values are: String, Number, Date, Boolean.

```typescript
type?: boolean;
```

Property Value
- boolean

Remarks
- [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### value

Specifies the value of the custom property. Note that even though Word on the web and the docx file format allow these properties to be arbitrarily long, the desktop version of Word will truncate string values to 255 16-bit chars (possibly creating invalid unicode by breaking up a surrogate pair).

```typescript
value?: boolean;
```

Property Value
- boolean

Remarks
- [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)