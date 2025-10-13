# Word.CustomXmlNode class

Package: https://learn.microsoft.com/en-us/javascript/api/word

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents an XML node in a tree in the document. The CustomXmlNode object is a member of the Word.CustomXmlNodeCollection object: https://learn.microsoft.com/en-us/javascript/api/word/word.customxmlnodecollection

- Extends: OfficeExtension.ClientObject (https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientobject)

## Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

## Properties

- attributes: Gets a CustomXmlNodeCollection object representing the attributes of the current element in the current node.
- baseName: Gets the base name of the node without the namespace prefix, if one exists.
- childNodes: Gets a CustomXmlNodeCollection object containing all of the child elements of the current node.
- context: The request context associated with the object. This connects the add-in's process to the Office host application's process.
- firstChild: Gets a CustomXmlNode object corresponding to the first child element of the current node. If the node has no child elements (or if it isn't of type CustomXmlNodeType.element), returns `Nothing`.
- lastChild: Gets a CustomXmlNode object corresponding to the last child element of the current node. If the node has no child elements (or if it isn't of type CustomXmlNodeType.element), the property returns `Nothing`.
- namespaceUri: Gets the unique address identifier for the namespace of the node.
- nextSibling: Gets the next sibling node (element, comment, or processing instruction) of the current node. If the node is the last sibling at its level, the property returns `Nothing`.
- nodeType: Gets the type of the current node.
- nodeValue: Specifies the value of the current node.
- ownerPart: Gets the object representing the part associated with this node.
- parentNode: Gets the parent element node of the current node. If the current node is at the root level, the property returns `Nothing`.
- previousSibling: Gets the previous sibling node (element, comment, or processing instruction) of the current node. If the current node is the first sibling at its level, the property returns `Nothing`.
- text: Specifies the text for the current node.
- xml: Gets the XML representation of the current node and its children.
- xpath: Gets a string with the canonicalized XPath for the current node. If the node is no longer in the Document Object Model (DOM), the property returns an error message.

## Methods

- appendChildNode(options): Appends a single node as the last child under the context element node in the tree.
- appendChildSubtree(xml): Adds a subtree as the last child under the context element node in the tree.
- delete(): Deletes the current node from the tree (including all of its children, if any exist).
- hasChildNodes(): Specifies if the current element node has child element nodes.
- insertNodeBefore(options): Inserts a new node just before the context node in the tree.
- insertSubtreeBefore(xml, options): Inserts the specified subtree into the location just before the context node.
- load(options): Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- load(propertyNames): Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- load(propertyNamesAndPaths): Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- removeChild(child): Removes the specified child node from the tree.
- replaceChildNode(oldNode, options): Removes the specified child node and replaces it with a different node in the same location.
- replaceChildSubtree(xml, oldNode): Removes the specified node and replaces it with a different subtree in the same location.
- selectNodes(xPath): Selects a collection of nodes matching an XPath expression.
- selectSingleNode(xPath): Selects a single node from a collection matching an XPath expression.
- set(properties, options): Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
- set(properties): Sets multiple properties on the object at the same time, based on an existing loaded object.
- toJSON(): Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.CustomXmlNode` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.CustomXmlNodeData`) that contains shallow copies of any loaded child properties from the original object.
- track(): Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
- untrack(): Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

## Property Details

### attributes

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets a `CustomXmlNodeCollection` object representing the attributes of the current element in the current node.

```typescript
readonly attributes: Word.CustomXmlNodeCollection;
```

- Property Value: Word.CustomXmlNodeCollection (https://learn.microsoft.com/en-us/javascript/api/word/word.customxmlnodecollection)

Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### baseName

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the base name of the node without the namespace prefix, if one exists.

```typescript
readonly baseName: string;
```

- Property Value: string

Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### childNodes

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets a `CustomXmlNodeCollection` object containing all of the child elements of the current node.

```typescript
readonly childNodes: Word.CustomXmlNodeCollection;
```

- Property Value: Word.CustomXmlNodeCollection (https://learn.microsoft.com/en-us/javascript/api/word/word.customxmlnodecollection)

Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### context

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

- Property Value: Word.RequestContext (https://learn.microsoft.com/en-us/javascript/api/word/word.requestcontext)

### firstChild

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets a `CustomXmlNode` object corresponding to the first child element of the current node. If the node has no child elements (or if it isn't of type CustomXmlNodeType.element), returns `Nothing`.

```typescript
readonly firstChild: Word.CustomXmlNode;
```

- Property Value: Word.CustomXmlNode (https://learn.microsoft.com/en-us/javascript/api/word/word.customxmlnode)

Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### lastChild

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets a `CustomXmlNode` object corresponding to the last child element of the current node. If the node has no child elements (or if it isn't of type CustomXmlNodeType.element), the property returns `Nothing`.

```typescript
readonly lastChild: Word.CustomXmlNode;
```

- Property Value: Word.CustomXmlNode (https://learn.microsoft.com/en-us/javascript/api/word/word.customxmlnode)

Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### namespaceUri

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the unique address identifier for the namespace of the node.

```typescript
readonly namespaceUri: string;
```

- Property Value: string

Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### nextSibling

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the next sibling node (element, comment, or processing instruction) of the current node. If the node is the last sibling at its level, the property returns `Nothing`.

```typescript
readonly nextSibling: Word.CustomXmlNode;
```

- Property Value: Word.CustomXmlNode (https://learn.microsoft.com/en-us/javascript/api/word/word.customxmlnode)

Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### nodeType

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the type of the current node.

```typescript
readonly nodeType: Word.CustomXmlNodeType | "element" | "attribute" | "text" | "cData" | "processingInstruction" | "comment" | "document";
```

- Property Value: Word.CustomXmlNodeType (https://learn.microsoft.com/en-us/javascript/api/word/word.customxmlnodetype) | "element" | "attribute" | "text" | "cData" | "processingInstruction" | "comment" | "document"

Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### nodeValue

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the value of the current node.

```typescript
nodeValue: string;
```

- Property Value: string

Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### ownerPart

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the object representing the part associated with this node.

```typescript
readonly ownerPart: Word.CustomXmlPart;
```

- Property Value: Word.CustomXmlPart (https://learn.microsoft.com/en-us/javascript/api/word/word.customxmlpart)

Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### parentNode

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the parent element node of the current node. If the current node is at the root level, the property returns `Nothing`.

```typescript
readonly parentNode: Word.CustomXmlNode;
```

- Property Value: Word.CustomXmlNode (https://learn.microsoft.com/en-us/javascript/api/word/word.customxmlnode)

Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### previousSibling

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the previous sibling node (element, comment, or processing instruction) of the current node. If the current node is the first sibling at its level, the property returns `Nothing`.

```typescript
readonly previousSibling: Word.CustomXmlNode;
```

- Property Value: Word.CustomXmlNode (https://learn.microsoft.com/en-us/javascript/api/word/word.customxmlnode)

Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### text

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the text for the current node.

```typescript
text: string;
```

- Property Value: string

Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### xml

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the XML representation of the current node and its children.

```typescript
readonly xml: string;
```

- Property Value: string

Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### xpath

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets a string with the canonicalized XPath for the current node. If the node is no longer in the Document Object Model (DOM), the property returns an error message.

```typescript
readonly xpath: string;
```

- Property Value: string

Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

## Method Details

### appendChildNode(options)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Appends a single node as the last child under the context element node in the tree.

```typescript
appendChildNode(options?: Word.CustomXmlAppendChildNodeOptions): OfficeExtension.ClientResult<number>;
```

- Parameters:
  - options: Word.CustomXmlAppendChildNodeOptions (https://learn.microsoft.com/en-us/javascript/api/word/word.customxmlappendchildnodeoptions)

  Optional. The options that define the node to be appended.

- Returns: OfficeExtension.ClientResult<number> (https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientresult)

Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### appendChildSubtree(xml)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Adds a subtree as the last child under the context element node in the tree.

```typescript
appendChildSubtree(xml: string): OfficeExtension.ClientResult<number>;
```

- Parameters:
  - xml: string

  A string representing the XML subtree.

- Returns: OfficeExtension.ClientResult<number> (https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientresult)

Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### delete()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Deletes the current node from the tree (including all of its children, if any exist).

```typescript
delete(): void;
```

- Returns: void

Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### hasChildNodes()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the current element node has child element nodes.

```typescript
hasChildNodes(): OfficeExtension.ClientResult<boolean>;
```

- Returns: OfficeExtension.ClientResult<boolean> (https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientresult)

Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### insertNodeBefore(options)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Inserts a new node just before the context node in the tree.

```typescript
insertNodeBefore(options?: Word.CustomXmlInsertNodeBeforeOptions): OfficeExtension.ClientResult<number>;
```

- Parameters:
  - options: Word.CustomXmlInsertNodeBeforeOptions (https://learn.microsoft.com/en-us/javascript/api/word/word.customxmlinsertnodebeforeoptions)

  Optional. The options that define the node to be inserted.

- Returns: OfficeExtension.ClientResult<number> (https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientresult)

Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### insertSubtreeBefore(xml, options)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Inserts the specified subtree into the location just before the context node.

```typescript
insertSubtreeBefore(xml: string, options?: Word.CustomXmlInsertSubtreeBeforeOptions): OfficeExtension.ClientResult<number>;
```

- Parameters:
  - xml: string

  A string representing the XML subtree.

  - options: Word.CustomXmlInsertSubtreeBeforeOptions (https://learn.microsoft.com/en-us/javascript/api/word/word.customxmlinsertsubtreebeforeoptions)

  Optional. The options available for inserting the subtree.

- Returns: OfficeExtension.ClientResult<number> (https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientresult)

Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### load(options)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(options?: Word.Interfaces.CustomXmlNodeLoadOptions): Word.CustomXmlNode;
```

- Parameters:
  - options: Word.Interfaces.CustomXmlNodeLoadOptions (https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.customxmlnodeloadoptions)

  Provides options for which properties of the object to load.

- Returns: Word.CustomXmlNode (https://learn.microsoft.com/en-us/javascript/api/word/word.customxmlnode)

### load(propertyNames)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.CustomXmlNode;
```

- Parameters:
  - propertyNames: string | string[]

  A comma-delimited string or an array of strings that specify the properties to load.

- Returns: Word.CustomXmlNode (https://learn.microsoft.com/en-us/javascript/api/word/word.customxmlnode)

### load(propertyNamesAndPaths)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.CustomXmlNode;
```

- Parameters:
  - propertyNamesAndPaths:
    {
    select?: string;
    expand?: string;
    }

  `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

- Returns: Word.CustomXmlNode (https://learn.microsoft.com/en-us/javascript/api/word/word.customxmlnode)

### removeChild(child)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Removes the specified child node from the tree.

```typescript
removeChild(child: Word.CustomXmlNode): OfficeExtension.ClientResult<number>;
```

- Parameters:
  - child: Word.CustomXmlNode (https://learn.microsoft.com/en-us/javascript/api/word/word.customxmlnode)

  The child node to remove.

- Returns: OfficeExtension.ClientResult<number> (https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientresult)

Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### replaceChildNode(oldNode, options)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Removes the specified child node and replaces it with a different node in the same location.

```typescript
replaceChildNode(oldNode: Word.CustomXmlNode, options?: Word.CustomXmlReplaceChildNodeOptions): OfficeExtension.ClientResult<number>;
```

- Parameters:
  - oldNode: Word.CustomXmlNode (https://learn.microsoft.com/en-us/javascript/api/word/word.customxmlnode)

  The node to be replaced.

  - options: Word.CustomXmlReplaceChildNodeOptions (https://learn.microsoft.com/en-us/javascript/api/word/word.customxmlreplacechildnodeoptions)

  Optional. The options that define the child node which is to replace the old node.

- Returns: OfficeExtension.ClientResult<number> (https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientresult)

Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### replaceChildSubtree(xml, oldNode)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Removes the specified node and replaces it with a different subtree in the same location.

```typescript
replaceChildSubtree(xml: string, oldNode: Word.CustomXmlNode): OfficeExtension.ClientResult<number>;
```

- Parameters:
  - xml: string

  A string representing the new subtree.

  - oldNode: Word.CustomXmlNode (https://learn.microsoft.com/en-us/javascript/api/word/word.customxmlnode)

  The node to be replaced.

- Returns: OfficeExtension.ClientResult<number> (https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientresult)

Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### selectNodes(xPath)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Selects a collection of nodes matching an XPath expression.

```typescript
selectNodes(xPath: string): Word.CustomXmlNodeCollection;
```

- Parameters:
  - xPath: string

  The XPath expression.

- Returns: Word.CustomXmlNodeCollection (https://learn.microsoft.com/en-us/javascript/api/word/word.customxmlnodecollection)

Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### selectSingleNode(xPath)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Selects a single node from a collection matching an XPath expression.

```typescript
selectSingleNode(xPath: string): Word.CustomXmlNode;
```

- Parameters:
  - xPath: string

  The XPath expression.

- Returns: Word.CustomXmlNode (https://learn.microsoft.com/en-us/javascript/api/word/word.customxmlnode)

Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### set(properties, options)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.CustomXmlNodeUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

- Parameters:
  - properties: Word.Interfaces.CustomXmlNodeUpdateData (https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.customxmlnodeupdatedata)

  A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.

  - options: OfficeExtension.UpdateOptions (https://learn.microsoft.com/en-us/javascript/api/office/officeextension.updateoptions)

  Provides an option to suppress errors if the properties object tries to set any read-only properties.

- Returns: void

### set(properties)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.CustomXmlNode): void;
```

- Parameters:
  - properties: Word.CustomXmlNode (https://learn.microsoft.com/en-us/javascript/api/word/word.customxmlnode)

- Returns: void

### toJSON()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.CustomXmlNode` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.CustomXmlNodeData`) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.CustomXmlNodeData;
```

- Returns: Word.Interfaces.CustomXmlNodeData (https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.customxmlnodedata)

### track()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject): https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member. If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.CustomXmlNode;
```

- Returns: Word.CustomXmlNode (https://learn.microsoft.com/en-us/javascript/api/word/word.customxmlnode)

### untrack()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject): https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

```typescript
untrack(): Word.CustomXmlNode;
```

- Returns: Word.CustomXmlNode (https://learn.microsoft.com/en-us/javascript/api/word/word.customxmlnode)