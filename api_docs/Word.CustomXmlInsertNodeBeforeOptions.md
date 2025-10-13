# Word.CustomXmlInsertNodeBeforeOptions interface

Package: [word](/en-us/javascript/api/word)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Inserts a new node just before the context node in the tree.

## Remarks

[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- name: If provided, specifies the base name of the element to be inserted.
- namespaceUri: If provided, specifies the namespace of the element to be inserted. This property is required to insert nodes of [type](/en-us/javascript/api/word/word.customxmlnodetype) `element` or `attribute`; otherwise, it's ignored.
- nextSibling: If provided, specifies the context node.
- nodeType: If provided, specifies the type of node to append. If the property isn't specified, it's assumed to be of type `element`.
- nodeValue: If provided, specifies the value of the inserted node for those nodes that allow text. If the node doesn't allow text, the property is ignored.

## Property Details

### name

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

If provided, specifies the base name of the element to be inserted.

```typescript
name?: string;
```

#### Property Value
string

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### namespaceUri

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

If provided, specifies the namespace of the element to be inserted. This property is required to insert nodes of [type](/en-us/javascript/api/word/word.customxmlnodetype) `element` or `attribute`; otherwise, it's ignored.

```typescript
namespaceUri?: string;
```

#### Property Value
string

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### nextSibling

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

If provided, specifies the context node.

```typescript
nextSibling?: Word.CustomXmlNode;
```

#### Property Value
[Word.CustomXmlNode](/en-us/javascript/api/word/word.customxmlnode)

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### nodeType

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

If provided, specifies the type of node to append. If the property isn't specified, it's assumed to be of type `element`.

```typescript
nodeType?: Word.CustomXmlNodeType | "element" | "attribute" | "text" | "cData" | "processingInstruction" | "comment" | "document";
```

#### Property Value
[Word.CustomXmlNodeType](/en-us/javascript/api/word/word.customxmlnodetype) | "element" | "attribute" | "text" | "cData" | "processingInstruction" | "comment" | "document"

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### nodeValue

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

If provided, specifies the value of the inserted node for those nodes that allow text. If the node doesn't allow text, the property is ignored.

```typescript
nodeValue?: string;
```

#### Property Value
string

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)