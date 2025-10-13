# Word.Interfaces.CustomPropertyData interface

Package: [word](/en-us/javascript/api/word)

An interface describing the data returned by calling `customProperty.toJSON()`.

## Properties

- key: Gets the key of the custom property.
- type: Gets the value type of the custom property. Possible values are: String, Number, Date, Boolean.
- value: Specifies the value of the custom property. Note that even though Word on the web and the docx file format allow these properties to be arbitrarily long, the desktop version of Word will truncate string values to 255 16-bit chars (possibly creating invalid unicode by breaking up a surrogate pair).

## Property Details

### key

Gets the key of the custom property.

```typescript
key?: string;
```

Property Value: string

Remarks: [ [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

### type

Gets the value type of the custom property. Possible values are: String, Number, Date, Boolean.

```typescript
type?: Word.DocumentPropertyType | "String" | "Number" | "Date" | "Boolean";
```

Property Value: [Word.DocumentPropertyType](/en-us/javascript/api/word/word.documentpropertytype) | "String" | "Number" | "Date" | "Boolean"

Remarks: [ [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

### value

Specifies the value of the custom property. Note that even though Word on the web and the docx file format allow these properties to be arbitrarily long, the desktop version of Word will truncate string values to 255 16-bit chars (possibly creating invalid unicode by breaking up a surrogate pair).

```typescript
value?: any;
```

Property Value: any

Remarks: [ [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]