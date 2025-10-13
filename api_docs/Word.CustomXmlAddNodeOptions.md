# Word.CustomXmlAddNodeOptions interface

Package: [word](/en-us/javascript/api/word)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

The options for adding a node to the XML tree.

## Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- name  
  If provided, specifies the base name of the element to be added.

- namespaceUri  
  If provided, specifies the namespace of the element to be appended. This property is required to add nodes of type element or attribute; otherwise, it's ignored.

- nextSibling  
  If provided, specifies the node which should become the next sibling of the new node. If not specified, the node is added to the end of the parent node's children. This property is ignored for additions of type attribute. If the node isn't a child of the parent, an error is displayed.

- nodeType  
  If provided, specifies the type of node to add. If the parameter isn't specified, it's assumed to be of type element.

- nodeValue  
  If provided, specifies the value of the added node for those nodes that allow text. If the node doesn't allow text, the property is ignored.

## Property Details

### name

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

If provided, specifies the base name of the element to be added.

```
name?: string;
```

Property Value  
string

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### namespaceUri

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

If provided, specifies the namespace of the element to be appended. This property is required to add nodes of type element or attribute; otherwise, it's ignored.

```
namespaceUri?: string;
```

Property Value  
string

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### nextSibling

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

If provided, specifies the node which should become the next sibling of the new node. If not specified, the node is added to the end of the parent node's children. This property is ignored for additions of type attribute. If the node isn't a child of the parent, an error is displayed.

```
nextSibling?: Word.CustomXmlNode;
```

Property Value  
[Word.CustomXmlNode](/en-us/javascript/api/word/word.customxmlnode)

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### nodeType

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

If provided, specifies the type of node to add. If the parameter isn't specified, it's assumed to be of type element.

```
nodeType?: Word.CustomXmlNodeType | "element" | "attribute" | "text" | "cData" | "processingInstruction" | "comment" | "document";
```

Property Value  
[Word.CustomXmlNodeType](/en-us/javascript/api/word/word.customxmlnodetype) | "element" | "attribute" | "text" | "cData" | "processingInstruction" | "comment" | "document"

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### nodeValue

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

If provided, specifies the value of the added node for those nodes that allow text. If the node doesn't allow text, the property is ignored.

```
nodeValue?: string;
```

Property Value  
string

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)