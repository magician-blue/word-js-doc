# Word.Interfaces.CustomPropertyUpdateData interface

Package: [word](/en-us/javascript/api/word)

An interface for updating data on the CustomProperty object, for use in customProperty.set({ ... }).

## Properties

- [value](#value): Specifies the value of the custom property. Note that even though Word on the web and the docx file format allow these properties to be arbitrarily long, the desktop version of Word will truncate string values to 255 16-bit chars (possibly creating invalid unicode by breaking up a surrogate pair).

## Property Details

### value

Specifies the value of the custom property. Note that even though Word on the web and the docx file format allow these properties to be arbitrarily long, the desktop version of Word will truncate string values to 255 16-bit chars (possibly creating invalid unicode by breaking up a surrogate pair).

```typescript
value?: any;
```

Property value: any

Remarks

- [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)