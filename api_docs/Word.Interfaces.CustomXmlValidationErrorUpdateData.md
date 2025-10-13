# Word.Interfaces.CustomXmlValidationErrorUpdateData interface

Package: [word](/en-us/javascript/api/word)

An interface for updating data on the `CustomXmlValidationError` object, for use in `customXmlValidationError.set({ ... })`.

## Properties

- node: Gets the node associated with this `CustomXmlValidationError` object, if any exist.If no nodes exist, the property returns `Nothing`.

## Property Details

### node

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the node associated with this `CustomXmlValidationError` object, if any exist.If no nodes exist, the property returns `Nothing`.

```typescript
node?: Word.Interfaces.CustomXmlNodeUpdateData;
```

#### Property Value

[Word.Interfaces.CustomXmlNodeUpdateData](/en-us/javascript/api/word/word.interfaces.customxmlnodeupdatedata)

#### Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)