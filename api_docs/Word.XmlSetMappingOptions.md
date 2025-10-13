# Word.XmlSetMappingOptions interface

Package: [word](/en-us/javascript/api/word)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

The options that define the prefix mapping and the source of the custom XML data.

## Remarks

[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- prefixMapping  
  If provided, specifies the prefix mappings to use when querying the expression provided in the `xPath` parameter of the `XmlMapping.setMapping` calling method. If omitted, Word uses the set of prefix mappings for the specified custom XML part in the current document.

- source  
  If provided, specifies the desired custom XML data to map the content control to. If this property is omitted, the XPath is evaluated against all custom XML in the current document, and the mapping is established with the first `CustomXmlPart` where the XPath resolves to an XML node.

## Property Details

### prefixMapping

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

If provided, specifies the prefix mappings to use when querying the expression provided in the `xPath` parameter of the `XmlMapping.setMapping` calling method. If omitted, Word uses the set of prefix mappings for the specified custom XML part in the current document.

```typescript
prefixMapping?: string;
```

Property Value: string

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### source

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

If provided, specifies the desired custom XML data to map the content control to. If this property is omitted, the XPath is evaluated against all custom XML in the current document, and the mapping is established with the first `CustomXmlPart` where the XPath resolves to an XML node.

```typescript
source?: Word.CustomXmlPart;
```

Property Value: [Word.CustomXmlPart](/en-us/javascript/api/word/word.customxmlpart)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)