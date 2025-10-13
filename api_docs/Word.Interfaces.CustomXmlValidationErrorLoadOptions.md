# Word.Interfaces.CustomXmlValidationErrorLoadOptions interface

- Package: [word](/en-us/javascript/api/word)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents a single validation error in a [Word.CustomXmlValidationErrorCollection](/en-us/javascript/api/word/word.customxmlvalidationerrorcollection) object.

## Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- $all  
  Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).

- errorCode  
  Gets an integer representing the validation error in the CustomXmlValidationError object.

- name  
  Gets the name of the error in the CustomXmlValidationError object.If no errors exist, the property returns Nothing

- node  
  Gets the node associated with this CustomXmlValidationError object, if any exist.If no nodes exist, the property returns Nothing.

- text  
  Gets the text in the CustomXmlValidationError object.

- type  
  Gets the type of error generated from the CustomXmlValidationError object.

## Property Details

### $all

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).

```typescript
$all?: boolean;
```

Property Value: boolean

### errorCode

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets an integer representing the validation error in the CustomXmlValidationError object.

```typescript
errorCode?: boolean;
```

Property Value: boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### name

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the name of the error in the CustomXmlValidationError object.If no errors exist, the property returns Nothing

```typescript
name?: boolean;
```

Property Value: boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### node

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the node associated with this CustomXmlValidationError object, if any exist.If no nodes exist, the property returns Nothing.

```typescript
node?: Word.Interfaces.CustomXmlNodeLoadOptions;
```

Property Value: [Word.Interfaces.CustomXmlNodeLoadOptions](/en-us/javascript/api/word/word.interfaces.customxmlnodeloadoptions)

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### text

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the text in the CustomXmlValidationError object.

```typescript
text?: boolean;
```

Property Value: boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### type

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the type of error generated from the CustomXmlValidationError object.

```typescript
type?: boolean;
```

Property Value: boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)