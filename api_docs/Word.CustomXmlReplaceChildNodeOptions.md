# Word.CustomXmlReplaceChildNodeOptions interface

Package: [word](/en-us/javascript/api/word)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Removes the specified child node and replaces it with a different node in the same location.

## Remarks

[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- [name](#name): If provided, specifies the base name of the replacement element.
- [namespaceUri](#namespaceuri): If provided, specifies the namespace of the replacement element. This property is required to replace nodes of [type](/en-us/javascript/api/word/word.customxmlnodetype) `element` or `attribute`; otherwise, it's ignored.
- [nodeType](#nodetype): If provided, specifies the type of the replacement node. If the property isn't specified, it's assumed to be of type `element`.
- [nodeValue](#nodevalue): If provided, specifies the value of the replacement node for those nodes that allow text. If the node doesn't allow text, the property is ignored.

## Property Details

### name

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

If provided, specifies the base name of the replacement element.

```typescript
name?: string;
```

#### Property Value
string

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### namespaceUri

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

If provided, specifies the namespace of the replacement element. This property is required to replace nodes of [type](/en-us/javascript/api/word/word.customxmlnodetype) `element` or `attribute`; otherwise, it's ignored.

```typescript
namespaceUri?: string;
```

#### Property Value
string

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### nodeType

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

If provided, specifies the type of the replacement node. If the property isn't specified, it's assumed to be of type `element`.

```typescript
nodeType?: Word.CustomXmlNodeType | "element" | "attribute" | "text" | "cData" | "processingInstruction" | "comment" | "document";
```

#### Property Value
[Word.CustomXmlNodeType](/en-us/javascript/api/word/word.customxmlnodetype) | "element" | "attribute" | "text" | "cData" | "processingInstruction" | "comment" | "document"

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### nodeValue

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

If provided, specifies the value of the replacement node for those nodes that allow text. If the node doesn't allow text, the property is ignored.

```typescript
nodeValue?: string;
```

#### Property Value
string

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)