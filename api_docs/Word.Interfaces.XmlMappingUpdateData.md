# Word.Interfaces.XmlMappingUpdateData interface

Package: [word](https://learn.microsoft.com/en-us/javascript/api/word)

An interface for updating data on the `XmlMapping` object, for use in `xmlMapping.set({ ... })`.

## Properties

- `customXmlNode`  
  Returns a `CustomXmlNode` object that represents the custom XML node in the data store that the content control in the document maps to.
- `customXmlPart`  
  Returns a `CustomXmlPart` object that represents the custom XML part to which the content control in the document maps.

## Property Details

### customXmlNode

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `CustomXmlNode` object that represents the custom XML node in the data store that the content control in the document maps to.

```typescript
customXmlNode?: Word.Interfaces.CustomXmlNodeUpdateData;
```

Property Value: [Word.Interfaces.CustomXmlNodeUpdateData](https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.customxmlnodeupdatedata)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### customXmlPart

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `CustomXmlPart` object that represents the custom XML part to which the content control in the document maps.

```typescript
customXmlPart?: Word.Interfaces.CustomXmlPartUpdateData;
```

Property Value: [Word.Interfaces.CustomXmlPartUpdateData](https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.customxmlpartupdatedata)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)