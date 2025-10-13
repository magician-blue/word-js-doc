# Word.Interfaces.CustomXmlNodeCollectionLoadOptions interface

- Package: [word](/en-us/javascript/api/word)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Contains a collection of [Word.CustomXmlNode](/en-us/javascript/api/word/word.customxmlnode) objects representing the XML nodes in a document.

## Remarks

[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- $all  
  Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

- baseName  
  For EACH ITEM in the collection: Gets the base name of the node without the namespace prefix, if one exists.

- firstChild  
  For EACH ITEM in the collection: Gets a `CustomXmlNode` object corresponding to the first child element of the current node. If the node has no child elements (or if it isn't of type [CustomXmlNodeType.element](/en-us/javascript/api/word/word.customxmlnodetype)), returns `Nothing`.

- lastChild  
  For EACH ITEM in the collection: Gets a `CustomXmlNode` object corresponding to the last child element of the current node. If the node has no child elements (or if it isn't of type [CustomXmlNodeType.element](/en-us/javascript/api/word/word.customxmlnodetype)), the property returns `Nothing`.

- namespaceUri  
  For EACH ITEM in the collection: Gets the unique address identifier for the namespace of the node.

- nextSibling  
  For EACH ITEM in the collection: Gets the next sibling node (element, comment, or processing instruction) of the current node. If the node is the last sibling at its level, the property returns `Nothing`.

- nodeType  
  For EACH ITEM in the collection: Gets the type of the current node.

- nodeValue  
  For EACH ITEM in the collection: Specifies the value of the current node.

- ownerPart  
  For EACH ITEM in the collection: Gets the object representing the part associated with this node.

- parentNode  
  For EACH ITEM in the collection: Gets the parent element node of the current node. If the current node is at the root level, the property returns `Nothing`.

- previousSibling  
  For EACH ITEM in the collection: Gets the previous sibling node (element, comment, or processing instruction) of the current node. If the current node is the first sibling at its level, the property returns `Nothing`.

- text  
  For EACH ITEM in the collection: Specifies the text for the current node.

- xml  
  For EACH ITEM in the collection: Gets the XML representation of the current node and its children.

- xpath  
  For EACH ITEM in the collection: Gets a string with the canonicalized XPath for the current node. If the node is no longer in the Document Object Model (DOM), the property returns an error message.

## Property Details

### $all

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

Property Value  
boolean

### baseName

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Gets the base name of the node without the namespace prefix, if one exists.

```typescript
baseName?: boolean;
```

Property Value  
boolean

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### firstChild

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Gets a `CustomXmlNode` object corresponding to the first child element of the current node. If the node has no child elements (or if it isn't of type [CustomXmlNodeType.element](/en-us/javascript/api/word/word.customxmlnodetype)), returns `Nothing`.

```typescript
firstChild?: Word.Interfaces.CustomXmlNodeLoadOptions;
```

Property Value  
[Word.Interfaces.CustomXmlNodeLoadOptions](/en-us/javascript/api/word/word.interfaces.customxmlnodeloadoptions)

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### lastChild

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Gets a `CustomXmlNode` object corresponding to the last child element of the current node. If the node has no child elements (or if it isn't of type [CustomXmlNodeType.element](/en-us/javascript/api/word/word.customxmlnodetype)), the property returns `Nothing`.

```typescript
lastChild?: Word.Interfaces.CustomXmlNodeLoadOptions;
```

Property Value  
[Word.Interfaces.CustomXmlNodeLoadOptions](/en-us/javascript/api/word/word.interfaces.customxmlnodeloadoptions)

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### namespaceUri

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Gets the unique address identifier for the namespace of the node.

```typescript
namespaceUri?: boolean;
```

Property Value  
boolean

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### nextSibling

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Gets the next sibling node (element, comment, or processing instruction) of the current node. If the node is the last sibling at its level, the property returns `Nothing`.

```typescript
nextSibling?: Word.Interfaces.CustomXmlNodeLoadOptions;
```

Property Value  
[Word.Interfaces.CustomXmlNodeLoadOptions](/en-us/javascript/api/word/word.interfaces.customxmlnodeloadoptions)

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### nodeType

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Gets the type of the current node.

```typescript
nodeType?: boolean;
```

Property Value  
boolean

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### nodeValue

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies the value of the current node.

```typescript
nodeValue?: boolean;
```

Property Value  
boolean

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### ownerPart

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Gets the object representing the part associated with this node.

```typescript
ownerPart?: Word.Interfaces.CustomXmlPartLoadOptions;
```

Property Value  
[Word.Interfaces.CustomXmlPartLoadOptions](/en-us/javascript/api/word/word.interfaces.customxmlpartloadoptions)

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### parentNode

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Gets the parent element node of the current node. If the current node is at the root level, the property returns `Nothing`.

```typescript
parentNode?: Word.Interfaces.CustomXmlNodeLoadOptions;
```

Property Value  
[Word.Interfaces.CustomXmlNodeLoadOptions](/en-us/javascript/api/word/word.interfaces.customxmlnodeloadoptions)

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### previousSibling

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Gets the previous sibling node (element, comment, or processing instruction) of the current node. If the current node is the first sibling at its level, the property returns `Nothing`.

```typescript
previousSibling?: Word.Interfaces.CustomXmlNodeLoadOptions;
```

Property Value  
[Word.Interfaces.CustomXmlNodeLoadOptions](/en-us/javascript/api/word/word.interfaces.customxmlnodeloadoptions)

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### text

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies the text for the current node.

```typescript
text?: boolean;
```

Property Value  
boolean

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### xml

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Gets the XML representation of the current node and its children.

```typescript
xml?: boolean;
```

Property Value  
boolean

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### xpath

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Gets a string with the canonicalized XPath for the current node. If the node is no longer in the Document Object Model (DOM), the property returns an error message.

```typescript
xpath?: boolean;
```

Property Value  
boolean

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)