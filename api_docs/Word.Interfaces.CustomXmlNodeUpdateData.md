# Word.Interfaces.CustomXmlNodeUpdateData interface

- Package: [word](/en-us/javascript/api/word)

An interface for updating data on the `CustomXmlNode` object, for use in `customXmlNode.set({ ... })`.

## Properties

- firstChild: Gets a `CustomXmlNode` object corresponding to the first child element of the current node. If the node has no child elements (or if it isn't of type [CustomXmlNodeType.element](/en-us/javascript/api/word/word.customxmlnodetype)), returns `Nothing`.
- lastChild: Gets a `CustomXmlNode` object corresponding to the last child element of the current node. If the node has no child elements (or if it isn't of type [CustomXmlNodeType.element](/en-us/javascript/api/word/word.customxmlnodetype)), the property returns `Nothing`.
- nextSibling: Gets the next sibling node (element, comment, or processing instruction) of the current node. If the node is the last sibling at its level, the property returns `Nothing`.
- nodeValue: Specifies the value of the current node.
- ownerPart: Gets the object representing the part associated with this node.
- parentNode: Gets the parent element node of the current node. If the current node is at the root level, the property returns `Nothing`.
- previousSibling: Gets the previous sibling node (element, comment, or processing instruction) of the current node. If the current node is the first sibling at its level, the property returns `Nothing`.
- text: Specifies the text for the current node.

## Property Details

### firstChild

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets a `CustomXmlNode` object corresponding to the first child element of the current node. If the node has no child elements (or if it isn't of type [CustomXmlNodeType.element](/en-us/javascript/api/word/word.customxmlnodetype)), returns `Nothing`.

```typescript
firstChild?: Word.Interfaces.CustomXmlNodeUpdateData;
```

Property Value: [Word.Interfaces.CustomXmlNodeUpdateData](/en-us/javascript/api/word/word.interfaces.customxmlnodeupdatedata)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### lastChild

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets a `CustomXmlNode` object corresponding to the last child element of the current node. If the node has no child elements (or if it isn't of type [CustomXmlNodeType.element](/en-us/javascript/api/word/word.customxmlnodetype)), the property returns `Nothing`.

```typescript
lastChild?: Word.Interfaces.CustomXmlNodeUpdateData;
```

Property Value: [Word.Interfaces.CustomXmlNodeUpdateData](/en-us/javascript/api/word/word.interfaces.customxmlnodeupdatedata)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### nextSibling

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the next sibling node (element, comment, or processing instruction) of the current node. If the node is the last sibling at its level, the property returns `Nothing`.

```typescript
nextSibling?: Word.Interfaces.CustomXmlNodeUpdateData;
```

Property Value: [Word.Interfaces.CustomXmlNodeUpdateData](/en-us/javascript/api/word/word.interfaces.customxmlnodeupdatedata)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### nodeValue

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the value of the current node.

```typescript
nodeValue?: string;
```

Property Value: string

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### ownerPart

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the object representing the part associated with this node.

```typescript
ownerPart?: Word.Interfaces.CustomXmlPartUpdateData;
```

Property Value: [Word.Interfaces.CustomXmlPartUpdateData](/en-us/javascript/api/word/word.interfaces.customxmlpartupdatedata)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### parentNode

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the parent element node of the current node. If the current node is at the root level, the property returns `Nothing`.

```typescript
parentNode?: Word.Interfaces.CustomXmlNodeUpdateData;
```

Property Value: [Word.Interfaces.CustomXmlNodeUpdateData](/en-us/javascript/api/word/word.interfaces.customxmlnodeupdatedata)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### previousSibling

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the previous sibling node (element, comment, or processing instruction) of the current node. If the current node is the first sibling at its level, the property returns `Nothing`.

```typescript
previousSibling?: Word.Interfaces.CustomXmlNodeUpdateData;
```

Property Value: [Word.Interfaces.CustomXmlNodeUpdateData](/en-us/javascript/api/word/word.interfaces.customxmlnodeupdatedata)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### text

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the text for the current node.

```typescript
text?: string;
```

Property Value: string

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)