# Word.Interfaces.CustomXmlValidationErrorCollectionLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents a collection of [Word.CustomXmlValidationError](/en-us/javascript/api/word/word.customxmlvalidationerror) objects.

## Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

## Properties

- $all  
  Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).
- errorCode  
  For EACH ITEM in the collection: Gets an integer representing the validation error in the CustomXmlValidationError object.
- name  
  For EACH ITEM in the collection: Gets the name of the error in the CustomXmlValidationError object.If no errors exist, the property returns Nothing
- node  
  For EACH ITEM in the collection: Gets the node associated with this CustomXmlValidationError object, if any exist.If no nodes exist, the property returns Nothing.
- text  
  For EACH ITEM in the collection: Gets the text in the CustomXmlValidationError object.
- type  
  For EACH ITEM in the collection: Gets the type of error generated from the CustomXmlValidationError object.

## Property Details

### $all

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).

```typescript
$all?: boolean;
```

Property Value  
boolean

### errorCode

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Gets an integer representing the validation error in the CustomXmlValidationError object.

```typescript
errorCode?: boolean;
```

Property Value  
boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

### name

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Gets the name of the error in the CustomXmlValidationError object.If no errors exist, the property returns Nothing

```typescript
name?: boolean;
```

Property Value  
boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

### node

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Gets the node associated with this CustomXmlValidationError object, if any exist.If no nodes exist, the property returns Nothing.

```typescript
node?: Word.Interfaces.CustomXmlNodeLoadOptions;
```

Property Value  
[Word.Interfaces.CustomXmlNodeLoadOptions](/en-us/javascript/api/word/word.interfaces.customxmlnodeloadoptions)

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

### text

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Gets the text in the CustomXmlValidationError object.

```typescript
text?: boolean;
```

Property Value  
boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

### type

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Gets the type of error generated from the CustomXmlValidationError object.

```typescript
type?: boolean;
```

Property Value  
boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]