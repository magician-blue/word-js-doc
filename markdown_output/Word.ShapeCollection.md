# Word.ShapeCollection

**Package:** `word`

**API Set:** WordApiDesktop 1.2

**Extends:** `OfficeExtension.ClientObject`

## Description

Contains a collection of [Word.Shape](/en-us/javascript/api/word/word.shape) objects. Currently, only the following shapes are supported: text boxes, geometric shapes, groups, pictures, and canvases.

## Class Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-shapes-text-boxes.yaml

await Word.run(async (context) => {
  // Gets text boxes in main document.
  const shapes: Word.ShapeCollection = context.document.body.shapes;
  shapes.load();
  await context.sync();

  if (shapes.items.length > 0) {
    shapes.items.forEach(function(shape, index) {
      if (shape.type === Word.ShapeType.textBox) {
        console.log(`Shape ${index} in the main document has a text box. Properties:`, shape);
      }
    });
  } else {
    console.log("No shapes found in main document.");
  }
});
```

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a ShapeCollection to verify the connection between the add-in and Word, then use it to load and sync shape properties.

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.getShapes();
    
    // Access the request context from the ShapeCollection
    const requestContext = shapes.context;
    
    // Use the context to load properties
    shapes.load("items");
    await requestContext.sync();
    
    console.log(`Found ${shapes.items.length} shapes in the document`);
    console.log(`Context is connected: ${requestContext !== null}`);
});
```

---

### items

**Type:** `Word.Shape[]`

Gets the loaded child items in this collection.

#### Examples

**Example**: Iterate through all shapes in the document and log their names and types to the console.

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.shapes;
    shapes.load("items");
    
    await context.sync();
    
    console.log(`Found ${shapes.items.length} shapes in the document`);
    
    for (let i = 0; i < shapes.items.length; i++) {
        const shape = shapes.items[i];
        shape.load("name, type");
    }
    
    await context.sync();
    
    shapes.items.forEach((shape, index) => {
        console.log(`Shape ${index + 1}: Name="${shape.name}", Type="${shape.type}"`);
    });
});
```

---

## Methods

### getByGeometricTypes

**Kind:** `read`

Gets the shapes that have the specified geometric types. Only applied to geometric shapes.

#### Signature

**Parameters:**
- `types`: `Word.GeometricShapeType[]` (required)
  An array of geometric shape subtypes.

**Returns:** `Word.ShapeCollection`

#### Examples

**Example**: Get all rectangle and circle shapes from the document and change their fill color to blue

```typescript
await Word.run(async (context) => {
    // Get all shapes in the document body
    const shapes = context.document.body.shapes;
    
    // Filter to get only rectangles and circles
    const geometricShapes = shapes.getByGeometricTypes([
        Word.GeometricShapeType.rectangle,
        Word.GeometricShapeType.oval
    ]);
    
    // Load the shapes
    geometricShapes.load("items");
    await context.sync();
    
    // Change fill color to blue for each shape
    for (let i = 0; i < geometricShapes.items.length; i++) {
        geometricShapes.items[i].fill.setSolidColor("blue");
    }
    
    await context.sync();
    
    console.log(`Found and updated ${geometricShapes.items.length} geometric shapes`);
});
```

---

### getById

**Kind:** `read`

Gets a shape by its identifier. Throws an `ItemNotFound` error if there isn't a shape with the identifier in this collection.

#### Signature

**Parameters:**
- `id`: `number` (required)
  A shape identifier.

**Returns:** `Word.Shape`

#### Examples

**Example**: Get a shape by its known identifier and change its fill color to blue

```typescript
await Word.run(async (context) => {
    // Assuming you have a known shape ID (e.g., from a previous operation)
    const shapeId = "12345678-1234-1234-1234-123456789012";
    
    // Get the shape by its ID
    const shape = context.document.body.shapes.getById(shapeId);
    
    // Change the shape's fill color to blue
    shape.fill.setSolidColor("blue");
    
    await context.sync();
    
    console.log("Shape found and updated successfully");
});
```

---

### getByIdOrNullObject

**Kind:** `read`

Gets a shape by its identifier. If there isn't a shape with the identifier in this collection, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

#### Signature

**Parameters:**
- `id`: `number` (required)
  A shape identifier.

**Returns:** `Word.Shape`

#### Examples

**Example**: Check if a shape with a specific ID exists in the document and display its type, or show a message if the shape is not found.

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.shapes;
    const shapeId = "12345"; // The ID of the shape to find
    
    const shape = shapes.getByIdOrNullObject(shapeId);
    shape.load("isNullObject, type");
    
    await context.sync();
    
    if (shape.isNullObject) {
        console.log(`Shape with ID ${shapeId} was not found.`);
    } else {
        console.log(`Shape found! Type: ${shape.type}`);
    }
});
```

---

### getByIds

**Kind:** `read`

Gets the shapes by the identifiers.

#### Signature

**Parameters:**
- `ids`: `number[]` (required)
  An array of shape identifiers.

**Returns:** `Word.ShapeCollection`

#### Examples

**Example**: Retrieve specific shapes by their IDs and change their fill color to light blue

```typescript
await Word.run(async (context) => {
    // Array of shape IDs to retrieve
    const shapeIds = ["shape1-id-123", "shape2-id-456", "shape3-id-789"];
    
    // Get the shapes by their IDs
    const shapes = context.document.body.shapes.getByIds(shapeIds);
    
    // Load the fill property for each shape
    shapes.load("items/fill");
    
    await context.sync();
    
    // Change the fill color of each retrieved shape
    for (let i = 0; i < shapes.items.length; i++) {
        shapes.items[i].fill.setSolidColor("#ADD8E6"); // Light blue
    }
    
    await context.sync();
});
```

---

### getByNames

**Kind:** `read`

Gets the shapes that have the specified names.

#### Signature

**Parameters:**
- `names`: `string[]` (required)
  An array of shape names.

**Returns:** `Word.ShapeCollection`

#### Examples

**Example**: Get shapes named "Logo" and "Signature" from the document and change their fill color to blue.

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.shapes;
    const targetShapes = shapes.getByNames(["Logo", "Signature"]);
    
    targetShapes.load("items");
    await context.sync();
    
    for (let i = 0; i < targetShapes.items.length; i++) {
        targetShapes.items[i].fill.setSolidColor("blue");
    }
    
    await context.sync();
});
```

---

### getByTypes

**Kind:** `read`

Gets the shapes that have the specified types.

#### Signature

**Parameters:**
- `types`: `Word.ShapeType[]` (required)
  An array of shape types.

**Returns:** `Word.ShapeCollection`

#### Examples

**Example**: Retrieve the first text box shape from the document body and update its position to top 115, left 0, and resize it to 50x50 dimensions.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-shapes-text-boxes.yaml

await Word.run(async (context) => {
  // Sets the properties of the first text box.
  const firstShapeWithTextBox: Word.Shape = context.document.body.shapes
    .getByTypes([Word.ShapeType.textBox])
    .getFirst();
  firstShapeWithTextBox.top = 115;
  firstShapeWithTextBox.left = 0;
  firstShapeWithTextBox.width = 50;
  firstShapeWithTextBox.height = 50;
  await context.sync();

  console.log("The first text box's properties were updated:", firstShapeWithTextBox);
});
```

---

### getFirst

**Kind:** `read`

Gets the first shape in this collection. Throws an `ItemNotFound` error if this collection is empty.

#### Signature

**Returns:** `Word.Shape`

#### Examples

**Example**: Insert a content control into the first paragraph of the first text box shape in the document.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-shapes-text-boxes.yaml

await Word.run(async (context) => {
  // Inserts a content control into the first paragraph in the first text box.
  const firstShapeWithTextBox: Word.Shape = context.document.body.shapes
    .getByTypes([Word.ShapeType.textBox])
    .getFirst();
  firstShapeWithTextBox.load("type/body");
  await context.sync();

  const firstParagraphInTextBox: Word.Paragraph = firstShapeWithTextBox.body.paragraphs.getFirst();
  const newControl: Word.ContentControl = firstParagraphInTextBox.insertContentControl();
  newControl.load();
  await context.sync();

  console.log("New content control properties:", newControl);
});
```

---

### getFirstOrNullObject

**Kind:** `read`

Gets the first shape in this collection. If this collection is empty, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

#### Signature

**Returns:** `Word.Shape`

#### Examples

**Example**: Check if a document contains any shapes and display the type of the first shape if it exists, or show a message if no shapes are present.

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.shapes;
    const firstShape = shapes.getFirstOrNullObject();
    firstShape.load("type, isNullObject");
    
    await context.sync();
    
    if (firstShape.isNullObject) {
        console.log("No shapes found in the document.");
    } else {
        console.log(`First shape type: ${firstShape.type}`);
    }
});
```

---

### group

**Kind:** `create`

Groups floating shapes in this collection, inline shapes will be skipped. Returns a Shape object that represents the new group of shapes.

#### Signature

**Returns:** `Word.Shape`

#### Examples

**Example**: Group all floating shapes in the document into a single group

```typescript
await Word.run(async (context) => {
    // Get all shapes in the document body
    const shapes = context.document.body.shapes;
    shapes.load("items");
    
    await context.sync();
    
    // Group all floating shapes (inline shapes will be automatically skipped)
    const shapeGroup = shapes.group();
    shapeGroup.load("id, name");
    
    await context.sync();
    
    console.log(`Created shape group with ID: ${shapeGroup.id}, Name: ${shapeGroup.name}`);
});
```

---

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.ShapeCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.ShapeCollection`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.ShapeCollection`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `OfficeExtension.LoadOption` (optional)
    `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.ShapeCollection`

#### Examples

**Example**: Load and display the count and types of all shapes in the document

```typescript
await Word.run(async (context) => {
    // Get the shapes collection from the document body
    const shapes = context.document.body.shapes;
    
    // Load the count and type properties of all shapes
    shapes.load("items/type");
    
    // Sync to execute the load command
    await context.sync();
    
    // Display the results
    console.log(`Total shapes in document: ${shapes.items.length}`);
    shapes.items.forEach((shape, index) => {
        console.log(`Shape ${index + 1}: ${shape.type}`);
    });
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.ShapeCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.ShapeCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

#### Signature

**Returns:** `Word.Interfaces.ShapeCollectionData`

#### Examples

**Example**: Get a JSON representation of all shapes in the document to log their properties for debugging purposes

```typescript
await Word.run(async (context) => {
    // Get all shapes in the document
    const shapes = context.document.body.shapes;
    
    // Load properties we want to serialize
    shapes.load("items/id, items/name, items/type, items/width, items/height");
    
    await context.sync();
    
    // Convert the ShapeCollection to a plain JavaScript object
    const shapesJSON = shapes.toJSON();
    
    // Log the JSON representation
    console.log(JSON.stringify(shapesJSON, null, 2));
    
    // You can now work with the plain object
    console.log(`Total shapes found: ${shapesJSON.items.length}`);
    shapesJSON.items.forEach(shape => {
        console.log(`Shape: ${shape.name}, Type: ${shape.type}, Size: ${shape.width}x${shape.height}`);
    });
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.ShapeCollection`

#### Examples

**Example**: Track a shape collection across multiple sync calls to maintain object references when modifying shape properties outside of a single batch operation

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    // Track the collection to use it across multiple sync calls
    shapes.track();
    
    // First sync - get initial data
    await context.sync();
    
    // Modify shapes in a subsequent operation
    for (let i = 0; i < shapes.items.length; i++) {
        shapes.items[i].fill.setSolidColor("blue");
    }
    
    await context.sync();
    
    // Untrack when done to free up memory
    shapes.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

#### Signature

**Returns:** `Word.ShapeCollection`

#### Examples

**Example**: Get all shapes in the document, process them to log their IDs, then untrack the collection to free memory

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.shapes;
    shapes.load("items");
    
    await context.sync();
    
    // Process the shapes
    console.log(`Found ${shapes.items.length} shapes`);
    shapes.items.forEach(shape => {
        console.log(`Shape ID: ${shape.id}`);
    });
    
    // Release memory associated with the shapes collection
    shapes.untrack();
    
    await context.sync();
});
```

---

## Source

- /en-us/javascript/api/word
- /en-us/javascript/api/office/officeextension.clientobject
- /en-us/javascript/api/requirement-sets/word/word-api-requirement-sets
- /en-us/javascript/api/word/word.shape
- /en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties
- /en-us/javascript/api/word/word.requestcontext
- /en-us/javascript/api/word/word.geometricshapetype
- /en-us/javascript/api/word/word.shapecollection
- /en-us/javascript/api/word/word.shape
- /en-us/javascript/api/word/word.shapetype
- /en-us/javascript/api/word/word.interfaces.shapecollectionloadoptions
- /en-us/javascript/api/word/word.interfaces.collectionloadoptions
- /en-us/javascript/api/office/officeextension.loadoption
- /en-us/javascript/api/word/word.interfaces.shapecollectiondata
- /en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member
