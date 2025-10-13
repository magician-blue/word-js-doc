# Word.Interfaces.CustomXmlPartUpdateData interface

Package: [word](/en-us/javascript/api/word)

An interface for updating data on the CustomXmlPart object, for use in customXmlPart.set({ ... }).

## Properties

- documentElement  
  Gets the root element of a bound region of data in the document. If the region is empty, the property returns `Nothing`.

## Property Details

### documentElement

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the root element of a bound region of data in the document. If the region is empty, the property returns `Nothing`.

```typescript
documentElement?: Word.Interfaces.CustomXmlNodeUpdateData;
```

Property Value: [Word.Interfaces.CustomXmlNodeUpdateData](/en-us/javascript/api/word/word.interfaces.customxmlnodeupdatedata)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)