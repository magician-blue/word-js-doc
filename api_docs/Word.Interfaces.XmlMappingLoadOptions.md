# Word.Interfaces.XmlMappingLoadOptions interface

Package: [word](/en-us/javascript/api/word)

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents the XML mapping on a [Word.ContentControl](/en-us/javascript/api/word/word.contentcontrol) object between custom XML and that content control. An XML mapping is a link between the text in a content control and an XML element in the custom XML data store for this document.

## Remarks

[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- `$all`  
  Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

- `customXmlNode`  
  Returns a `CustomXmlNode` object that represents the custom XML node in the data store that the content control in the document maps to.

- `customXmlPart`  
  Returns a `CustomXmlPart` object that represents the custom XML part to which the content control in the document maps.

- `isMapped`  
  Returns whether the content control in the document is mapped to an XML node in the document's XML data store.

- `prefixMappings`  
  Returns the prefix mappings used to evaluate the XPath for the current XML mapping.

- `xpath`  
  Returns the XPath for the XML mapping, which evaluates to the currently mapped XML node.

## Property Details

### $all

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

- Property Value: boolean

### customXmlNode

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `CustomXmlNode` object that represents the custom XML node in the data store that the content control in the document maps to.

```typescript
customXmlNode?: Word.Interfaces.CustomXmlNodeLoadOptions;
```

- Property Value: [Word.Interfaces.CustomXmlNodeLoadOptions](/en-us/javascript/api/word/word.interfaces.customxmlnodeloadoptions)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### customXmlPart

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `CustomXmlPart` object that represents the custom XML part to which the content control in the document maps.

```typescript
customXmlPart?: Word.Interfaces.CustomXmlPartLoadOptions;
```

- Property Value: [Word.Interfaces.CustomXmlPartLoadOptions](/en-us/javascript/api/word/word.interfaces.customxmlpartloadoptions)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### isMapped

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns whether the content control in the document is mapped to an XML node in the document's XML data store.

```typescript
isMapped?: boolean;
```

- Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### prefixMappings

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the prefix mappings used to evaluate the XPath for the current XML mapping.

```typescript
prefixMappings?: boolean;
```

- Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### xpath

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the XPath for the XML mapping, which evaluates to the currently mapped XML node.

```typescript
xpath?: boolean;
```

- Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)