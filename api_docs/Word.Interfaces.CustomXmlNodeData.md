# Word.Interfaces.CustomXmlNodeData interface

- Package: [word](/en-us/javascript/api/word)

An interface describing the data returned by calling `customXmlNode.toJSON()`.

## Properties

- [attributes](#attributes): Gets a `CustomXmlNodeCollection` object representing the attributes of the current element in the current node.
- [baseName](#basename): Gets the base name of the node without the namespace prefix, if one exists.
- [childNodes](#childnodes): Gets a `CustomXmlNodeCollection` object containing all of the child elements of the current node.
- [firstChild](#firstchild): Gets a `CustomXmlNode` object corresponding to the first child element of the current node. If the node has no child elements (or if it isn't of type [CustomXmlNodeType.element](/en-us/javascript/api/word/word.customxmlnodetype)), returns `Nothing`.
- [lastChild](#lastchild): Gets a `CustomXmlNode` object corresponding to the last child element of the current node. If the node has no child elements (or if it isn't of type [CustomXmlNodeType.element](/en-us/javascript/api/word/word.customxmlnodetype)), the property returns `Nothing`.
- [namespaceUri](#namespaceuri): Gets the unique address identifier for the namespace of the node.
- [nextSibling](#nextsibling): Gets the next sibling node (element, comment, or processing instruction) of the current node. If the node is the last sibling at its level, the property returns `Nothing`.
- [nodeType](#nodetype): Gets the type of the current node.
- [nodeValue](#nodevalue): Specifies the value of the current node.
- [ownerPart](#ownerpart): Gets the object representing the part associated with this node.
- [parentNode](#parentnode): Gets the parent element node of the current node. If the current node is at the root level, the property returns `Nothing`.
- [previousSibling](#previoussibling): Gets the previous sibling node (element, comment, or processing instruction) of the current node. If the current node is the first sibling at its level, the property returns `Nothing`.
- [text](#text): Specifies the text for the current node.
- [xml](#xml): Gets the XML representation of the current node and its children.
- [xpath](#xpath): Gets a string with the canonicalized XPath for the current node. If the node is no longer in the Document Object Model (DOM), the property returns an error message.

## Property Details

### attributes

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets a `CustomXmlNodeCollection` object representing the attributes of the current element in the current node.

```typescript
attributes?: Word.Interfaces.CustomXmlNodeData[];
```

#### Property Value
[Word.Interfaces.CustomXmlNodeData](/en-us/javascript/api/word/word.interfaces.customxmlnodedata)[]

#### Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### baseName

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the base name of the node without the namespace prefix, if one exists.

```typescript
baseName?: string;
```

#### Property Value
string

#### Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### childNodes

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets a `CustomXmlNodeCollection` object containing all of the child elements of the current node.

```typescript
childNodes?: Word.Interfaces.CustomXmlNodeData[];
```

#### Property Value
[Word.Interfaces.CustomXmlNodeData](/en-us/javascript/api/word/word.interfaces.customxmlnodedata)[]

#### Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### firstChild

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets a `CustomXmlNode` object corresponding to the first child element of the current node. If the node has no child elements (or if it isn't of type [CustomXmlNodeType.element](/en-us/javascript/api/word/word.customxmlnodetype)), returns `Nothing`.

```typescript
firstChild?: Word.Interfaces.CustomXmlNodeData;
```

#### Property Value
[Word.Interfaces.CustomXmlNodeData](/en-us/javascript/api/word/word.interfaces.customxmlnodedata)

#### Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### lastChild

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets a `CustomXmlNode` object corresponding to the last child element of the current node. If the node has no child elements (or if it isn't of type [CustomXmlNodeType.element](/en-us/javascript/api/word/word.customxmlnodetype)), the property returns `Nothing`.

```typescript
lastChild?: Word.Interfaces.CustomXmlNodeData;
```

#### Property Value
[Word.Interfaces.CustomXmlNodeData](/en-us/javascript/api/word/word.interfaces.customxmlnodedata)

#### Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### namespaceUri

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the unique address identifier for the namespace of the node.

```typescript
namespaceUri?: string;
```

#### Property Value
string

#### Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### nextSibling

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the next sibling node (element, comment, or processing instruction) of the current node. If the node is the last sibling at its level, the property returns `Nothing`.

```typescript
nextSibling?: Word.Interfaces.CustomXmlNodeData;
```

#### Property Value
[Word.Interfaces.CustomXmlNodeData](/en-us/javascript/api/word/word.interfaces.customxmlnodedata)

#### Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### nodeType

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the type of the current node.

```typescript
nodeType?: Word.CustomXmlNodeType | "element" | "attribute" | "text" | "cData" | "processingInstruction" | "comment" | "document";
```

#### Property Value
[Word.CustomXmlNodeType](/en-us/javascript/api/word/word.customxmlnodetype) | "element" | "attribute" | "text" | "cData" | "processingInstruction" | "comment" | "document"

#### Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### nodeValue

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the value of the current node.

```typescript
nodeValue?: string;
```

#### Property Value
string

#### Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### ownerPart

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the object representing the part associated with this node.

```typescript
ownerPart?: Word.Interfaces.CustomXmlPartData;
```

#### Property Value
[Word.Interfaces.CustomXmlPartData](/en-us/javascript/api/word/word.interfaces.customxmlpartdata)

#### Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### parentNode

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the parent element node of the current node. If the current node is at the root level, the property returns `Nothing`.

```typescript
parentNode?: Word.Interfaces.CustomXmlNodeData;
```

#### Property Value
[Word.Interfaces.CustomXmlNodeData](/en-us/javascript/api/word/word.interfaces.customxmlnodedata)

#### Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### previousSibling

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the previous sibling node (element, comment, or processing instruction) of the current node. If the current node is the first sibling at its level, the property returns `Nothing`.

```typescript
previousSibling?: Word.Interfaces.CustomXmlNodeData;
```

#### Property Value
[Word.Interfaces.CustomXmlNodeData](/en-us/javascript/api/word/word.interfaces.customxmlnodedata)

#### Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### text

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the text for the current node.

```typescript
text?: string;
```

#### Property Value
string

#### Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### xml

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the XML representation of the current node and its children.

```typescript
xml?: string;
```

#### Property Value
string

#### Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### xpath

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets a string with the canonicalized XPath for the current node. If the node is no longer in the Document Object Model (DOM), the property returns an error message.

```typescript
xpath?: string;
```

#### Property Value
string

#### Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)