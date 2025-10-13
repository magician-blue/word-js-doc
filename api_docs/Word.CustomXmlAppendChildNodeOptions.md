# Word.CustomXmlAppendChildNodeOptions interface

Package: [word](/en-us/javascript/api/word)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

The options that define the prefix mapping and the source of the custom XML data.

## Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

## Properties

- name  
  If provided, specifies the base name of the element to be appended.

- namespaceUri  
  If provided, specifies the namespace of the element to be appended. This property is required to append nodes of [type](/en-us/javascript/api/word/word.customxmlnodetype) element or attribute; otherwise, it's ignored.

- nodeType  
  If provided, specifies the type of node to append. If the property isn't specified, it's assumed to be of type element.

- nodeValue  
  If provided, specifies the value of the appended node for those nodes that allow text. If the node doesn't allow text, the property is ignored.

## Property Details

### name

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

If provided, specifies the base name of the element to be appended.

```typescript
name?: string;
```

#### Property Value
string

#### Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### namespaceUri

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

If provided, specifies the namespace of the element to be appended. This property is required to append nodes of [type](/en-us/javascript/api/word/word.customxmlnodetype) element or attribute; otherwise, it's ignored.

```typescript
namespaceUri?: string;
```

#### Property Value
string

#### Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### nodeType

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

If provided, specifies the type of node to append. If the property isn't specified, it's assumed to be of type element.

```typescript
nodeType?: Word.CustomXmlNodeType | "element" | "attribute" | "text" | "cData" | "processingInstruction" | "comment" | "document";
```

#### Property Value
[Word.CustomXmlNodeType](/en-us/javascript/api/word/word.customxmlnodetype) | "element" | "attribute" | "text" | "cData" | "processingInstruction" | "comment" | "document"

#### Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### nodeValue

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

If provided, specifies the value of the appended node for those nodes that allow text. If the node doesn't allow text, the property is ignored.

```typescript
nodeValue?: string;
```

#### Property Value
string

#### Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]