# Word.Interfaces.XmlMappingData interface

Package: [word](/en-us/javascript/api/word)

An interface describing the data returned by calling `xmlMapping.toJSON()`.

## Properties

- customXmlNode  
  Returns a `CustomXmlNode` object that represents the custom XML node in the data store that the content control in the document maps to.
- customXmlPart  
  Returns a `CustomXmlPart` object that represents the custom XML part to which the content control in the document maps.
- isMapped  
  Returns whether the content control in the document is mapped to an XML node in the document's XML data store.
- prefixMappings  
  Returns the prefix mappings used to evaluate the XPath for the current XML mapping.
- xpath  
  Returns the XPath for the XML mapping, which evaluates to the currently mapped XML node.

## Property Details

### customXmlNode

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `CustomXmlNode` object that represents the custom XML node in the data store that the content control in the document maps to.

```typescript
customXmlNode?: Word.Interfaces.CustomXmlNodeData;
```

Property Value  
[Word.Interfaces.CustomXmlNodeData](/en-us/javascript/api/word/word.interfaces.customxmlnodedata)

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### customXmlPart

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `CustomXmlPart` object that represents the custom XML part to which the content control in the document maps.

```typescript
customXmlPart?: Word.Interfaces.CustomXmlPartData;
```

Property Value  
[Word.Interfaces.CustomXmlPartData](/en-us/javascript/api/word/word.interfaces.customxmlpartdata)

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### isMapped

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns whether the content control in the document is mapped to an XML node in the document's XML data store.

```typescript
isMapped?: boolean;
```

Property Value  
boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### prefixMappings

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the prefix mappings used to evaluate the XPath for the current XML mapping.

```typescript
prefixMappings?: string;
```

Property Value  
string

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### xpath

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the XPath for the XML mapping, which evaluates to the currently mapped XML node.

```typescript
xpath?: string;
```

Property Value  
string

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)