# Word.CustomXmlNode

**Package:** `word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents an XML node in a tree in the document. The CustomXmlNode object is a member of the Word.CustomXmlNodeCollection object: https://learn.microsoft.com/en-us/javascript/api/word/word.customxmlnodecollection

## Properties

### attributes

**Type:** `Word.CustomXmlNodeCollection`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets a CustomXmlNodeCollection object representing the attributes of the current element in the current node.

#### Examples

**Example**: Read and display all attribute names and values from a custom XML element node in the document.

```typescript
await Word.run(async (context) => {
    // Get the custom XML parts collection
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();

    // Assume we have a custom XML part with an element that has attributes
    const xmlPart = customXmlParts.items[0];
    const xmlNodes = xmlPart.getXml();
    await context.sync();

    // Get a specific element node (e.g., root element)
    const elementNode = xmlPart.query("//*[@id]")[0]; // Get first element with 'id' attribute
    
    if (elementNode) {
        // Get the attributes collection of this element
        const attributes = elementNode.attributes;
        attributes.load("items");
        await context.sync();

        // Display all attributes
        console.log(`Element has ${attributes.items.length} attributes:`);
        attributes.items.forEach(attr => {
            attr.load("baseName, nodeValue");
        });
        await context.sync();

        attributes.items.forEach(attr => {
            console.log(`${attr.baseName} = "${attr.nodeValue}"`);
        });
    }
});
```

---

### baseName

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the base name of the node without the namespace prefix, if one exists.

#### Examples

**Example**: Get the base name of a custom XML node (without namespace prefix) and display it in the console.

```typescript
await Word.run(async (context) => {
    // Get the custom XML parts in the document
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();

    if (customXmlParts.items.length > 0) {
        // Get the first custom XML part
        const xmlPart = customXmlParts.items[0];
        const xmlNodes = xmlPart.getXml();
        await context.sync();

        // Query for a specific node (e.g., a node with a specific XPath)
        const nodes = xmlPart.query("//*");
        nodes.load("items");
        await context.sync();

        if (nodes.items.length > 0) {
            const node = nodes.items[0];
            node.load("baseName");
            await context.sync();

            // Display the base name without namespace prefix
            console.log("Node base name: " + node.baseName);
        }
    }
});
```

---

### childNodes

**Type:** `Word.CustomXmlNodeCollection`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets a CustomXmlNodeCollection object containing all of the child elements of the current node.

#### Examples

**Example**: Get all child nodes of a custom XML node and log their count to the console.

```typescript
await Word.run(async (context) => {
    // Get the custom XML parts collection
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();

    // Get the first custom XML part (assuming it exists)
    const customXmlPart = customXmlParts.items[0];
    
    // Get the root node
    const rootNode = customXmlPart.getXml();
    
    // Get the child nodes of the root
    const childNodes = rootNode.childNodes;
    childNodes.load("items");
    await context.sync();

    // Log the count of child nodes
    console.log(`Number of child nodes: ${childNodes.items.length}`);
    
    // Optionally, iterate through child nodes
    for (let i = 0; i < childNodes.items.length; i++) {
        const childNode = childNodes.items[i];
        childNode.load("baseName");
        await context.sync();
        console.log(`Child node ${i}: ${childNode.baseName}`);
    }
});
```

---

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a CustomXmlNode to verify the connection between the add-in and Word, then use it to load and read the node's properties.

```typescript
await Word.run(async (context) => {
    // Get custom XML parts and a specific node
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();
    
    if (customXmlParts.items.length > 0) {
        const xmlPart = customXmlParts.items[0];
        const xmlNodes = xmlPart.getXml();
        await context.sync();
        
        // Get the root node
        const rootNode = xmlPart.getNodes("/*")[0];
        
        // Access the context property from the CustomXmlNode
        const nodeContext = rootNode.context;
        
        // Use the context to load properties
        rootNode.load("baseName,nodeType");
        await nodeContext.sync();
        
        console.log(`Node name: ${rootNode.baseName}`);
        console.log(`Node type: ${rootNode.nodeType}`);
    }
});
```

---

### firstChild

**Type:** `Word.CustomXmlNode`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets a CustomXmlNode object corresponding to the first child element of the current node. If the node has no child elements (or if it isn't of type CustomXmlNodeType.element), returns `Nothing`.

#### Examples

**Example**: Get the first child element of a custom XML node and display its base name in the console.

```typescript
await Word.run(async (context) => {
    // Get the custom XML parts collection
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();

    // Assume we have at least one custom XML part
    if (customXmlParts.items.length > 0) {
        const xmlPart = customXmlParts.items[0];
        const rootNode = xmlPart.getXml();
        
        // Get the first child of the root node
        const firstChild = rootNode.firstChild;
        
        if (firstChild) {
            firstChild.load("baseName");
            await context.sync();
            
            console.log("First child element name: " + firstChild.baseName);
        } else {
            console.log("No child elements found");
        }
    }
});
```

---

### lastChild

**Type:** `Word.CustomXmlNode`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets a CustomXmlNode object corresponding to the last child element of the current node. If the node has no child elements (or if it isn't of type CustomXmlNodeType.element), the property returns `Nothing`.

#### Examples

**Example**: Get the last child element of a custom XML node and display its base name in the console.

```typescript
await Word.run(async (context) => {
    // Get the custom XML parts in the document
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();

    if (customXmlParts.items.length > 0) {
        // Get the first custom XML part
        const xmlPart = customXmlParts.items[0];
        const xmlRoot = xmlPart.getXml();
        await context.sync();

        // Get nodes from the XML part
        const xmlNodes = xmlPart.getNodes("*");
        xmlNodes.load("items");
        await context.sync();

        if (xmlNodes.items.length > 0) {
            // Get the first node that might have children
            const parentNode = xmlNodes.items[0];
            
            // Get the last child of this node
            const lastChild = parentNode.lastChild;
            
            if (lastChild) {
                lastChild.load("baseName");
                await context.sync();
                
                console.log("Last child element name:", lastChild.baseName);
            } else {
                console.log("No child elements found.");
            }
        }
    }
});
```

---

### namespaceUri

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the unique address identifier for the namespace of the node.

#### Examples

**Example**: Get and display the namespace URI of a custom XML node in the document.

```typescript
await Word.run(async (context) => {
    // Get the custom XML parts in the document
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();

    if (customXmlParts.items.length > 0) {
        // Get the first custom XML part
        const xmlPart = customXmlParts.items[0];
        const xmlNodes = xmlPart.getXml();
        await context.sync();

        // Query for a specific node (e.g., root node or specific element)
        const nodes = xmlPart.query("//*");
        nodes.load("items");
        await context.sync();

        if (nodes.items.length > 0) {
            const node = nodes.items[0];
            node.load("namespaceUri");
            await context.sync();

            // Display the namespace URI
            console.log("Namespace URI: " + node.namespaceUri);
        }
    }
});
```

---

### nextSibling

**Type:** `Word.CustomXmlNode`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the next sibling node (element, comment, or processing instruction) of the current node. If the node is the last sibling at its level, the property returns `Nothing`.

#### Examples

**Example**: Navigate through sibling XML nodes in a custom XML part and log each sibling's base name to the console.

```typescript
await Word.run(async (context) => {
    // Get the first custom XML part
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();

    if (customXmlParts.items.length > 0) {
        const xmlPart = customXmlParts.items[0];
        const xmlNodes = xmlPart.getXml();
        
        // Get the root node and its first child
        const rootNode = xmlPart.getXmlNodeByXPath("/");
        const firstChild = rootNode.getFirstChild();
        firstChild.load("baseName");
        await context.sync();

        // Navigate through siblings using nextSibling
        let currentNode = firstChild;
        currentNode.load("baseName");
        await context.sync();

        while (currentNode.isNullObject === false) {
            console.log("Sibling node: " + currentNode.baseName);
            
            // Move to next sibling
            currentNode = currentNode.nextSibling;
            currentNode.load("baseName, isNullObject");
            await context.sync();
        }
    }
});
```

---

### nodeType

**Type:** `Word.CustomXmlNodeType | "element" | "attribute" | "text" | "cData" | "processingInstruction" | "comment" | "document"`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the type of the current node.

#### Examples

**Example**: Check the type of each XML node in a custom XML part and log different messages based on whether it's an element, attribute, text, or other node type.

```typescript
await Word.run(async (context) => {
    // Get the first custom XML part
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();

    if (customXmlParts.items.length > 0) {
        const xmlPart = customXmlParts.items[0];
        const xmlNodes = xmlPart.query("//*");
        xmlNodes.load("items");
        await context.sync();

        // Check the type of each node
        for (let i = 0; i < xmlNodes.items.length; i++) {
            const node = xmlNodes.items[i];
            node.load("nodeType");
        }
        await context.sync();

        // Log messages based on node type
        for (let i = 0; i < xmlNodes.items.length; i++) {
            const node = xmlNodes.items[i];
            
            if (node.nodeType === Word.CustomXmlNodeType.element || node.nodeType === "element") {
                console.log(`Node ${i} is an element node`);
            } else if (node.nodeType === Word.CustomXmlNodeType.attribute || node.nodeType === "attribute") {
                console.log(`Node ${i} is an attribute node`);
            } else if (node.nodeType === Word.CustomXmlNodeType.text || node.nodeType === "text") {
                console.log(`Node ${i} is a text node`);
            } else {
                console.log(`Node ${i} is of type: ${node.nodeType}`);
            }
        }
    }
});
```

---

### nodeValue

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the value of the current node.

#### Examples

**Example**: Read and update the value of a custom XML node in a Word document

```typescript
await Word.run(async (context) => {
    // Get the custom XML parts collection
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();

    // Get the first custom XML part (assuming it exists)
    const customXmlPart = customXmlParts.items[0];
    
    // Get a specific node by XPath query
    const nodes = customXmlPart.query("/root/data");
    nodes.load("items");
    await context.sync();

    if (nodes.items.length > 0) {
        const node = nodes.items[0];
        
        // Load the current node value
        node.load("nodeValue");
        await context.sync();
        
        console.log("Current node value:", node.nodeValue);
        
        // Update the node value
        node.nodeValue = "Updated value";
        
        await context.sync();
        console.log("Node value updated successfully");
    }
});
```

---

### ownerPart

**Type:** `Word.CustomXmlPart`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the object representing the part associated with this node.

#### Examples

**Example**: Get the namespace URI of the custom XML part that owns a specific node in the document.

```typescript
await Word.run(async (context) => {
    // Get the first custom XML part in the document
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();

    if (customXmlParts.items.length > 0) {
        const xmlPart = customXmlParts.items[0];
        
        // Get the root node of the XML part
        const rootNode = xmlPart.getXml();
        const nodes = xmlPart.query("//*");
        nodes.load("items");
        await context.sync();

        if (nodes.items.length > 0) {
            const firstNode = nodes.items[0];
            
            // Get the owner part of this node
            const ownerPart = firstNode.ownerPart;
            ownerPart.load("namespaceUri");
            await context.sync();

            console.log("Owner part namespace URI: " + ownerPart.namespaceUri);
        }
    }
});
```

---

### parentNode

**Type:** `Word.CustomXmlNode`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the parent element node of the current node. If the current node is at the root level, the property returns `Nothing`.

#### Examples

**Example**: Navigate from a child XML node to its parent node and display the parent node's base name in the console.

```typescript
await Word.run(async (context) => {
    // Get custom XML parts in the document
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();

    if (customXmlParts.items.length > 0) {
        // Get the first custom XML part
        const xmlPart = customXmlParts.items[0];
        const xmlNodes = xmlPart.query("//*");
        xmlNodes.load("items");
        await context.sync();

        if (xmlNodes.items.length > 0) {
            // Get a child node
            const childNode = xmlNodes.items[0];
            
            // Access the parent node
            const parentNode = childNode.parentNode;
            
            if (parentNode) {
                parentNode.load("baseName");
                await context.sync();
                
                console.log("Parent node name: " + parentNode.baseName);
            } else {
                console.log("This node is at the root level and has no parent.");
            }
        }
    }
});
```

---

### previousSibling

**Type:** `Word.CustomXmlNode`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the previous sibling node (element, comment, or processing instruction) of the current node. If the current node is the first sibling at its level, the property returns `Nothing`.

#### Examples

**Example**: Navigate to a specific XML node and retrieve its previous sibling node to compare their content.

```typescript
await Word.run(async (context) => {
    // Get custom XML parts with a specific namespace
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();

    if (customXmlParts.items.length > 0) {
        const xmlPart = customXmlParts.items[0];
        const xmlNodes = xmlPart.query("//*");
        xmlNodes.load("items");
        await context.sync();

        if (xmlNodes.items.length > 1) {
            const currentNode = xmlNodes.items[1];
            const previousSibling = currentNode.previousSibling;
            
            if (previousSibling) {
                previousSibling.load("baseName, nodeType");
                await context.sync();
                
                console.log(`Previous sibling name: ${previousSibling.baseName}`);
                console.log(`Previous sibling type: ${previousSibling.nodeType}`);
            } else {
                console.log("No previous sibling found - this is the first node.");
            }
        }
    }
});
```

---

### text

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the text for the current node.

#### Examples

**Example**: Get the text content from a custom XML node and display it in the document

```typescript
await Word.run(async (context) => {
    // Get the custom XML parts collection
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();

    // Get the first custom XML part (assuming it exists)
    const customXmlPart = customXmlParts.items[0];
    
    // Get the root node
    const rootNode = customXmlPart.getXml();
    const nodes = customXmlPart.query("//*");
    nodes.load("items");
    await context.sync();

    // Get the first node and access its text property
    const firstNode = nodes.items[0];
    firstNode.load("text");
    await context.sync();

    // Insert the node's text into the document
    const range = context.document.body.insertParagraph(
        `Node text: ${firstNode.text}`,
        Word.InsertLocation.end
    );

    await context.sync();
});
```

---

### xml

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the XML representation of the current node and its children.

#### Examples

**Example**: Retrieve and log the XML content of a custom XML part's root node to the console.

```typescript
await Word.run(async (context) => {
    // Get the first custom XML part in the document
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();

    if (customXmlParts.items.length > 0) {
        // Get the root node of the first custom XML part
        const xmlPart = customXmlParts.items[0];
        const rootNode = xmlPart.getXml();
        
        // Access the xml property to get the XML representation
        const xmlNodes = xmlPart.query("//*");
        xmlNodes.load("items");
        await context.sync();

        if (xmlNodes.items.length > 0) {
            const node = xmlNodes.items[0];
            node.load("xml");
            await context.sync();
            
            // Log the XML content of the node
            console.log("Node XML content:", node.xml);
        }
    }
});
```

---

### xpath

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets a string with the canonicalized XPath for the current node. If the node is no longer in the Document Object Model (DOM), the property returns an error message.

#### Examples

**Example**: Retrieve and display the XPath location of a custom XML node within the document's custom XML parts

```typescript
await Word.run(async (context) => {
    // Get the first custom XML part in the document
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();

    if (customXmlParts.items.length > 0) {
        const xmlPart = customXmlParts.items[0];
        
        // Get the root node
        const rootNode = xmlPart.getXml();
        await context.sync();
        
        // Get all nodes in the XML part
        const nodes = xmlPart.query("//*");
        nodes.load("items");
        await context.sync();
        
        if (nodes.items.length > 0) {
            const firstNode = nodes.items[0];
            
            // Get the XPath of the node
            firstNode.load("xpath");
            await context.sync();
            
            console.log("Node XPath: " + firstNode.xpath);
        }
    }
});
```

---

## Methods

### appendChildNode

**Kind:** `create`

Appends a single node as the last child under the context element node in the tree.

#### Signature

**Parameters:**
- `options`: `Word.CustomXmlAppendChildNodeOptions` (optional)
  Optional. The options that define the node to be appended.

**Returns:** `OfficeExtension.ClientResult<number>`

#### Examples

**Example**: Add a new child element node named "Author" with text content "John Smith" to an existing CustomXmlPart's root node

```typescript
await Word.run(async (context) => {
    // Get the first custom XML part
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();
    
    const customXmlPart = customXmlParts.items[0];
    const rootNode = customXmlPart.getXml();
    
    // Get the root element node
    const nodes = customXmlPart.query("//*");
    nodes.load("items");
    await context.sync();
    
    const rootElement = nodes.items[0];
    
    // Append a new child node named "Author" with text content
    rootElement.appendChildNode("Author", "element", "John Smith");
    
    await context.sync();
    console.log("Child node 'Author' appended successfully");
});
```

---

### appendChildSubtree

**Kind:** `create`

Adds a subtree as the last child under the context element node in the tree.

#### Signature

**Parameters:**
- `xml`: `string` (required)
  A string representing the XML subtree.

**Returns:** `OfficeExtension.ClientResult<number>`

#### Examples

**Example**: Add a new contact entry with name and email as a subtree to an existing custom XML node in the document

```typescript
await Word.run(async (context) => {
    // Get the custom XML parts collection
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();

    // Get the first custom XML part (or find a specific one by namespace)
    const customXmlPart = customXmlParts.items[0];
    
    // Get the root node
    const rootNode = customXmlPart.getXml();
    context.load(rootNode);
    await context.sync();

    // Get a specific parent node where we want to add the subtree
    const nodes = customXmlPart.query("//contacts");
    context.load(nodes, "items");
    await context.sync();

    if (nodes.items.length > 0) {
        const contactsNode = nodes.items[0];
        
        // Define the XML subtree to append
        const newContactXml = "<contact><name>John Doe</name><email>john@example.com</email></contact>";
        
        // Append the subtree as a child
        contactsNode.appendChildSubtree(newContactXml);
        
        await context.sync();
        console.log("Contact subtree added successfully");
    }
});
```

---

### delete

**Kind:** `delete`

Deletes the current node from the tree (including all of its children, if any exist).

#### Signature

**Returns:** `void`

#### Examples

**Example**: Delete a specific custom XML node (and all its children) from the document's custom XML parts

```typescript
await Word.run(async (context) => {
    // Get the first custom XML part in the document
    const customXmlPart = context.document.customXmlParts.getByNamespace("http://example.com/mydata").getOnlyItem();
    
    // Get the root node
    const rootNode = customXmlPart.getXml();
    
    // Query for a specific node to delete (e.g., a node with tag "obsoleteData")
    const nodesToDelete = rootNode.getChildNodes("obsoleteData");
    
    // Load the collection
    nodesToDelete.load("items");
    await context.sync();
    
    // Delete the first matching node (and all its children)
    if (nodesToDelete.items.length > 0) {
        nodesToDelete.items[0].delete();
        await context.sync();
        
        console.log("Custom XML node deleted successfully");
    }
});
```

---

### hasChildNodes

**Kind:** `read`

Specifies if the current element node has child element nodes.

#### Signature

**Returns:** `OfficeExtension.ClientResult<boolean>`

#### Examples

**Example**: Check if a custom XML node has child nodes and display the result in the console

```typescript
await Word.run(async (context) => {
    // Get the custom XML parts collection
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();

    if (customXmlParts.items.length > 0) {
        // Get the first custom XML part
        const xmlPart = customXmlParts.items[0];
        const xmlNodes = xmlPart.getXml();
        
        // Get the root node
        const rootNode = xmlPart.getNodes("/*");
        rootNode.load("items");
        await context.sync();

        if (rootNode.items.length > 0) {
            const node = rootNode.items[0];
            
            // Check if the node has child nodes
            const hasChildren = node.hasChildNodes();
            await context.sync();

            console.log(`Node has child nodes: ${hasChildren.value}`);
        }
    }
});
```

---

### insertNodeBefore

**Kind:** `create`

Inserts a new node just before the context node in the tree.

#### Signature

**Parameters:**
- `options`: `Word.CustomXmlInsertNodeBeforeOptions` (optional)
  Optional. The options that define the node to be inserted.

**Returns:** `OfficeExtension.ClientResult<number>`

#### Examples

**Example**: Insert a new XML node with a "priority" element before an existing "task" node in a custom XML part

```typescript
await Word.run(async (context) => {
    // Get the custom XML parts collection
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();

    // Get the first custom XML part (or find a specific one by namespace)
    const customXmlPart = customXmlParts.items[0];
    
    // Get the root node and find the target node
    const rootNode = customXmlPart.getXml();
    const nodes = customXmlPart.query("//task");
    nodes.load("items");
    await context.sync();

    if (nodes.items.length > 0) {
        const taskNode = nodes.items[0];
        
        // Insert a new "priority" node before the "task" node
        taskNode.insertNodeBefore("<priority>High</priority>");
        
        await context.sync();
        console.log("New node inserted before the task node");
    }
});
```

---

### insertSubtreeBefore

**Kind:** `create`

Inserts the specified subtree into the location just before the context node.

#### Signature

**Parameters:**
- `xml`: `string` (required)
  A string representing the XML subtree.
- `options`: `Word.CustomXmlInsertSubtreeBeforeOptions` (optional)
  Optional. The options available for inserting the subtree.

**Returns:** `OfficeExtension.ClientResult<number>`

#### Examples

**Example**: Insert a new XML subtree containing contact information before an existing "customer" node in a custom XML part

```typescript
await Word.run(async (context) => {
    // Get the custom XML parts collection
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();

    // Get the first custom XML part (or find a specific one)
    const customXmlPart = customXmlParts.items[0];
    
    // Get all nodes with tag name "customer"
    const customerNodes = customXmlPart.getNodesByXPath("//customer");
    customerNodes.load("items");
    await context.sync();

    // Get the first customer node
    const targetNode = customerNodes.items[0];
    
    // Define the new subtree to insert (contact information)
    const newSubtree = '<contact><email>info@example.com</email><phone>555-1234</phone></contact>';
    
    // Insert the subtree before the customer node
    targetNode.insertSubtreeBefore(newSubtree);
    
    await context.sync();
    console.log("Contact subtree inserted before customer node");
});
```

---

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.CustomXmlNodeLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.CustomXmlNode`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.CustomXmlNode`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.CustomXmlNode`

#### Examples

**Example**: Load and display the properties of a custom XML node, including its base name, namespace URI, and node type.

```typescript
await Word.run(async (context) => {
    // Get the custom XML parts collection
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();

    // Get the first custom XML part (assuming it exists)
    const customXmlPart = customXmlParts.items[0];
    
    // Get the root node
    const rootNode = customXmlPart.getXml();
    const xmlNodes = customXmlPart.query("//*");
    xmlNodes.load("items");
    await context.sync();

    // Get the first node
    const customXmlNode = xmlNodes.items[0];
    
    // Load specific properties of the node
    customXmlNode.load("baseName, namespaceUri, nodeType");
    await context.sync();

    // Display the loaded properties
    console.log("Base Name: " + customXmlNode.baseName);
    console.log("Namespace URI: " + customXmlNode.namespaceUri);
    console.log("Node Type: " + customXmlNode.nodeType);
});
```

---

### removeChild

**Kind:** `delete`

Removes the specified child node from the tree.

#### Signature

**Parameters:**
- `child`: `Word.CustomXmlNode` (required)
  The child node to remove.

**Returns:** `OfficeExtension.ClientResult<number>`

#### Examples

**Example**: Remove a specific child node from a custom XML part's node tree in the document

```typescript
await Word.run(async (context) => {
    // Get the custom XML parts collection
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();

    // Get the first custom XML part (or find by namespace)
    const customXmlPart = customXmlParts.items[0];
    
    // Get the root node
    const rootNode = customXmlPart.getXml();
    const nodes = customXmlPart.query("/root/child");
    nodes.load("items");
    await context.sync();

    // Get parent and child nodes
    const parentNode = nodes.items[0];
    const childNodes = parentNode.getChildNodes();
    childNodes.load("items");
    await context.sync();

    // Remove the first child node from its parent
    if (childNodes.items.length > 0) {
        const childToRemove = childNodes.items[0];
        parentNode.removeChild(childToRemove);
        await context.sync();
        
        console.log("Child node removed successfully");
    }
});
```

---

### replaceChildNode

**Kind:** `write`

Removes the specified child node and replaces it with a different node in the same location.

#### Signature

**Parameters:**
- `oldNode`: `Word.CustomXmlNode` (required)
  The node to be replaced.
- `options`: `Word.CustomXmlReplaceChildNodeOptions` (optional)
  Optional. The options that define the child node which is to replace the old node.

**Returns:** `OfficeExtension.ClientResult<number>`

#### Examples

**Example**: Replace an outdated XML child node with a new node containing updated data in a custom XML part

```typescript
await Word.run(async (context) => {
    // Get the custom XML parts collection
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();

    // Get the first custom XML part (or find by namespace)
    const customXmlPart = customXmlParts.items[0];
    
    // Get the root node
    const rootNode = customXmlPart.getXml();
    context.load(rootNode);
    await context.sync();

    // Get child nodes
    const childNodes = rootNode.getChildNodes();
    childNodes.load("items");
    await context.sync();

    // Identify the old node to replace (e.g., first child)
    const oldNode = childNodes.items[0];
    
    // Create new XML content for replacement
    const newXmlContent = '<updatedElement>New Value</updatedElement>';
    
    // Replace the old child node with the new one
    rootNode.replaceChildNode(oldNode, { xmlContent: newXmlContent });
    
    await context.sync();
    console.log("Child node replaced successfully");
});
```

---

### replaceChildSubtree

**Kind:** `write`

Removes the specified node and replaces it with a different subtree in the same location.

#### Signature

**Parameters:**
- `xml`: `string` (required)
  A string representing the new subtree.
- `oldNode`: `Word.CustomXmlNode` (required)
  The node to be replaced.

**Returns:** `OfficeExtension.ClientResult<number>`

#### Examples

**Example**: Replace an outdated contact information node with a new contact details subtree in a custom XML part

```typescript
await Word.run(async (context) => {
    // Get the custom XML parts collection
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();

    // Get the first custom XML part (or find by namespace)
    const customXmlPart = customXmlParts.items[0];
    
    // Get all nodes with tag name "contact"
    const contactNodes = customXmlPart.getXmlNodesByTagName("contact");
    contactNodes.load("items");
    await context.sync();

    // Get the parent node containing the old contact node
    const parentNode = contactNodes.items[0].parentNode;
    const oldContactNode = contactNodes.items[0];
    
    // Define new XML subtree to replace the old contact
    const newContactXml = '<contact><name>Jane Smith</name><email>jane.smith@example.com</email><phone>555-0199</phone></contact>';
    
    // Replace the old contact node with the new contact subtree
    parentNode.replaceChildSubtree(newContactXml, oldContactNode);
    
    await context.sync();
    console.log("Contact node replaced successfully");
});
```

---

### selectNodes

**Kind:** `read`

Selects a collection of nodes matching an XPath expression.

#### Signature

**Parameters:**
- `xPath`: `string` (required)
  The XPath expression.

**Returns:** `Word.CustomXmlNodeCollection`

#### Examples

**Example**: Select all "author" nodes from a custom XML part using an XPath expression and log their count to the console.

```typescript
await Word.run(async (context) => {
    // Get the first custom XML part in the document
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();
    
    if (customXmlParts.items.length > 0) {
        const xmlPart = customXmlParts.items[0];
        const xmlRoot = xmlPart.getXml();
        await context.sync();
        
        // Get the root node of the custom XML part
        const rootNode = xmlPart.getOnlyChild();
        
        // Select all nodes matching the XPath expression
        const authorNodes = rootNode.selectNodes("//author");
        authorNodes.load("items");
        await context.sync();
        
        console.log(`Found ${authorNodes.items.length} author nodes`);
    }
});
```

---

### selectSingleNode

**Kind:** `read`

Selects a single node from a collection matching an XPath expression.

#### Signature

**Parameters:**
- `xPath`: `string` (required)
  The XPath expression.

**Returns:** `Word.CustomXmlNode`

#### Examples

**Example**: Select and retrieve the text content of a specific employee node by ID from a Custom XML part using an XPath expression

```typescript
await Word.run(async (context) => {
    // Get the custom XML parts collection
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();

    // Get the first custom XML part (assumes it contains employee data)
    const xmlPart = customXmlParts.items[0];
    
    // Get the root node
    const rootNode = xmlPart.getXml();
    await context.sync();
    
    // Select a single employee node with ID="E001" using XPath
    const employeeNode = rootNode.selectSingleNode("//employee[@id='E001']");
    employeeNode.load("baseName, nodeType, namespaceUri");
    await context.sync();
    
    console.log(`Found node: ${employeeNode.baseName}`);
    console.log(`Node type: ${employeeNode.nodeType}`);
});
```

---

### set

**Kind:** `write`

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `Interfaces.CustomXmlNodeUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.CustomXmlNode` (required)

  **Returns:** `void`

#### Examples

**Example**: Update multiple properties of a custom XML node at once, setting both its base XML content and loading specific properties for inspection.

```typescript
await Word.run(async (context) => {
    // Get the first custom XML part and its root node
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();
    
    const customXmlPart = customXmlParts.items[0];
    const xmlNodes = customXmlPart.getNodes("//*");
    xmlNodes.load("items");
    await context.sync();
    
    const xmlNode = xmlNodes.items[0];
    
    // Use set() to configure multiple properties at once
    xmlNode.set({
        nodeValue: "Updated node value",
        // Additional properties can be set here
    });
    
    await context.sync();
    
    console.log("Custom XML node properties updated successfully");
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.CustomXmlNode` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.CustomXmlNodeData`) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.CustomXmlNodeData`

#### Examples

**Example**: Serialize a custom XML node to a plain JavaScript object and log its properties to the console for debugging purposes.

```typescript
await Word.run(async (context) => {
    // Get the custom XML parts in the document
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();

    if (customXmlParts.items.length > 0) {
        // Get the first custom XML part
        const xmlPart = customXmlParts.items[0];
        const xmlNodes = xmlPart.getXml();
        await context.sync();

        // Query for a specific node (e.g., first child node)
        const nodes = xmlPart.query("//*[1]");
        nodes.load("items");
        await context.sync();

        if (nodes.items.length > 0) {
            const node = nodes.items[0];
            node.load("baseName,namespaceUri,nodeType");
            await context.sync();

            // Convert the CustomXmlNode to a plain JavaScript object
            const nodeData = node.toJSON();
            
            // Log the serialized object for debugging
            console.log("Node data:", nodeData);
            console.log("Base name:", nodeData.baseName);
            console.log("Namespace URI:", nodeData.namespaceUri);
            console.log("Node type:", nodeData.nodeType);
        }
    }
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.CustomXmlNode`

#### Examples

**Example**: Track a custom XML node across multiple sync calls to safely access and modify its properties without encountering InvalidObjectPath errors

```typescript
await Word.run(async (context) => {
    // Get a custom XML part and find a specific node
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();
    
    const customXmlPart = customXmlParts.items[0];
    const xmlNodes = customXmlPart.getXml();
    await context.sync();
    
    // Get the root node
    const rootNode = customXmlPart.getNodes("/*")[0];
    rootNode.load("baseName");
    
    // Track the node to use it across multiple sync calls
    rootNode.track();
    
    await context.sync();
    
    // Now we can safely use the tracked node in subsequent operations
    console.log("Root node name: " + rootNode.baseName);
    
    // Perform additional operations with the tracked node
    const childNodes = rootNode.getChildNodes();
    childNodes.load("items");
    await context.sync();
    
    console.log("Number of child nodes: " + childNodes.items.length);
    
    // Untrack when done to release memory
    rootNode.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

#### Signature

**Returns:** `Word.CustomXmlNode`

#### Examples

**Example**: Retrieve a custom XML node, use it to read data, then untrack it to release memory after processing is complete.

```typescript
await Word.run(async (context) => {
    // Get custom XML parts and find a specific node
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();

    if (customXmlParts.items.length > 0) {
        const xmlPart = customXmlParts.items[0];
        const xmlNodes = xmlPart.query("//*");
        xmlNodes.load("items");
        await context.sync();

        if (xmlNodes.items.length > 0) {
            const node = xmlNodes.items[0];
            
            // Track the node to work with it
            context.trackedObjects.add(node);
            node.load("baseName, nodeType");
            await context.sync();

            // Use the node data
            console.log(`Node name: ${node.baseName}, Type: ${node.nodeType}`);

            // Untrack the node to free memory when done
            node.untrack();
            await context.sync();
        }
    }
});
```

---

## Source

- https://learn.microsoft.com/en-us/javascript/api/word/word.customxmlnode
