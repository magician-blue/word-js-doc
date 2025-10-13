# Word.Interfaces.BodyUpdateData interface

Package: https://learn.microsoft.com/en-us/javascript/api/word

An interface for updating data on the Body object, for use in body.set({ ... }).

## Properties

- font
  - Gets the text format of the body. Use this to get and set font name, size, color, and other properties.
- style
  - Specifies the style name for the body. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
- styleBuiltIn
  - Specifies the built-in style name for the body. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.

## Property Details

### font

Gets the text format of the body. Use this to get and set font name, size, color, and other properties.

```typescript
font?: Word.Interfaces.FontUpdateData;
```

Property Value
- Word.Interfaces.FontUpdateData: https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.fontupdatedata

Remarks
- [API set: WordApi 1.1](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### style

Specifies the style name for the body. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.

```typescript
style?: string;
```

Property Value
- string

Remarks
- [API set: WordApi 1.1](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### styleBuiltIn

Specifies the built-in style name for the body. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.

```typescript
styleBuiltIn?: Word.BuiltInStyleName | "Other" | "Normal" | "Heading1" | "Heading2" | "Heading3" | "Heading4" | "Headi