# Word.Interfaces.CustomXmlValidationErrorData interface

Package: https://learn.microsoft.com/en-us/javascript/api/word

An interface describing the data returned by calling `customXmlValidationError.toJSON()`.

## Properties

- errorCode  
  Gets an integer representing the validation error in the `CustomXmlValidationError` object.

- name  
  Gets the name of the error in the `CustomXmlValidationError` object.If no errors exist, the property returns `Nothing`

- node  
  Gets the node associated with this `CustomXmlValidationError` object, if any exist.If no nodes exist, the property returns `Nothing`.

- text  
  Gets the text in the `CustomXmlValidationError` object.

- type  
  Gets the type of error generated from the `CustomXmlValidationError` object.

## Property Details

### errorCode

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets an integer representing the validation error in the `CustomXmlValidationError` object.

```typescript
errorCode?: number;
```

#### Property Value
number

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### name

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the name of the error in the `CustomXmlValidationError` object.If no errors exist, the property returns `Nothing`

```typescript
name?: string;
```

#### Property Value
string

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### node

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the node associated with this `CustomXmlValidationError` object, if any exist.If no nodes exist, the property returns `Nothing`.

```typescript
node?: Word.Interfaces.CustomXmlNodeData;
```

#### Property Value
Word.Interfaces.CustomXmlNodeData  
https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.customxmlnodedata

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### text

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the text in the `CustomXmlValidationError` object.

```typescript
text?: string;
```

#### Property Value
string

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### type

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the type of error generated from the `CustomXmlValidationError` object.

```typescript
type?: Word.CustomXmlValidationErrorType | "schemaGenerated" | "automaticallyCleared" | "manual";
```

#### Property Value
[Word.CustomXmlValidationErrorType](https://learn.microsoft.com/en-us/javascript/api/word/word.customxmlvalidationerrortype) | "schemaGenerated" | "automaticallyCleared" | "manual"

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)