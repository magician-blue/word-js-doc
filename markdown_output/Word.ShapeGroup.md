# ShapeGroup

**Package:** `word`

**API Set:** WordApiDesktop 1.2

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents a shape group in the document. To get the corresponding Shape object, use ShapeGroup.shape.

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a ShapeGroup object to verify the connection between the add-in and Word application before performing operations on the shape group.

```typescript
await Word.run(async (context) => {
    // Get the first shape group in the document
    const shapeGroups = context.document.body.shapeGroups;
    shapeGroups.load("items");
    await context.sync();
    
    if (shapeGroups.items.length > 0) {
        const shapeGroup = shapeGroups.items[0];
        
        // Access the request context from the shape group
        const requestContext = shapeGroup.context;
        
        // Verify the context is valid by using it to load properties
        shapeGroup.load("id,name");
        await requestContext.sync();
        
        console.log(`Shape group accessed via context - ID: ${shapeGroup.id}, Name: ${shapeGroup.name}`);
    }
});
```

---

### id

**Type:** `number`

**Since:** WordApiDesktop 1.2

Gets an integer that represents the shape group identifier.

#### Examples

**Example**: Get the shape group identifier and display it in the console to track which shape group is being processed.

```typescript
await Word.run(async (context) => {
    const shapeGroups = context.document.body.shapeGroups;
    shapeGroups.load("id");
    
    await context.sync();
    
    if (shapeGroups.items.length > 0) {
        const shapeGroup = shapeGroups.items[0];
        console.log(`Shape Group ID: ${shapeGroup.id}`);
    }
});
```

---

### shape

**Type:** `Word.Shape`

**Since:** WordApiDesktop 1.2

Gets the Shape object associated with the group.

#### Examples

**Example**: Get the shape group's associated Shape object and change its fill color to blue.

```typescript
await Word.run(async (context) => {
    const shapeGroups = context.document.body.shapeGroups;
    shapeGroups.load("items");
    await context.sync();
    
    if (shapeGroups.items.length > 0) {
        const shapeGroup = shapeGroups.items[0];
        
        // Get the Shape object associated with the group
        const groupShape = shapeGroup.shape;
        groupShape.fill.setSolidColor("blue");
        
        await context.sync();
    }
});
```

---

### shapes

**Type:** `Word.ShapeCollection`

**Since:** WordApiDesktop 1.2

Gets the collection of Shape objects. Currently, only text boxes, geometric shapes, and pictures are supported.

#### Examples

**Example**: Get all shapes within a shape group and log their names to the console

```typescript
await Word.run(async (context) => {
    // Get the first shape group in the document
    const shapeGroups = context.document.body.shapeGroups;
    shapeGroups.load("items");
    await context.sync();
    
    if (shapeGroups.items.length > 0) {
        const shapeGroup = shapeGroups.items[0];
        
        // Get the shapes collection from the shape group
        const shapes = shapeGroup.shapes;
        shapes.load("items/name");
        await context.sync();
        
        // Log each shape's name
        shapes.items.forEach(shape => {
            console.log(`Shape name: ${shape.name}`);
        });
    }
});
```

---

## Methods

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.ShapeGroupLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.ShapeGroup`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.ShapeGroup`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.ShapeGroup`

#### Examples

**Example**: Load and display the ID and child shape count of the first shape group in the document.

```typescript
await Word.run(async (context) => {
    // Get the first shape group in the document
    const shapeGroup = context.document.body.shapeGroups.getFirst();
    
    // Load specific properties of the shape group
    shapeGroup.load("id, childShapeCount");
    
    // Sync to execute the load command
    await context.sync();
    
    // Display the loaded properties
    console.log(`Shape Group ID: ${shapeGroup.id}`);
    console.log(`Number of child shapes: ${shapeGroup.childShapeCount}`);
});
```

---

### set

**Kind:** `configure`

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `Interfaces.ShapeGroupUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.ShapeGroup` (required)

  **Returns:** `void`

#### Examples

**Example**: Update multiple properties of a shape group at once, including its name and position

```typescript
await Word.run(async (context) => {
    const shapeGroup = context.document.body.shapes.getItem(0).getAsShapeGroup();
    
    shapeGroup.set({
        name: "UpdatedShapeGroup",
        left: 100,
        top: 150
    });
    
    await context.sync();
    console.log("Shape group properties updated successfully");
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.ShapeGroup` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.ShapeGroupData`) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.ShapeGroupData`

#### Examples

**Example**: Serialize a shape group's properties to JSON format for logging or data export purposes.

```typescript
await Word.run(async (context) => {
    // Get the first shape group in the document
    const shapeGroups = context.document.body.shapeGroups;
    shapeGroups.load("items");
    await context.sync();

    if (shapeGroups.items.length > 0) {
        const shapeGroup = shapeGroups.items[0];
        
        // Load properties you want to serialize
        shapeGroup.load("id,name,width,height");
        await context.sync();

        // Convert the shape group to a plain JavaScript object
        const shapeGroupData = shapeGroup.toJSON();
        
        // Now you can use the plain object for logging, storage, etc.
        console.log("Shape Group Data:", JSON.stringify(shapeGroupData, null, 2));
    }
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.ShapeGroup`

#### Examples

**Example**: Track a shape group object to maintain its reference across multiple sync calls when modifying its properties

```typescript
await Word.run(async (context) => {
    // Get the first shape group in the document
    const shapeGroups = context.document.body.shapeGroups;
    shapeGroups.load("items");
    await context.sync();
    
    if (shapeGroups.items.length > 0) {
        const shapeGroup = shapeGroups.items[0];
        
        // Track the shape group to use it across multiple sync calls
        shapeGroup.track();
        
        // First sync - load properties
        shapeGroup.load("id,width,height");
        await context.sync();
        
        console.log(`Shape Group ID: ${shapeGroup.id}`);
        
        // Second sync - modify properties
        shapeGroup.left = 100;
        shapeGroup.top = 100;
        await context.sync();
        
        // Untrack when done
        shapeGroup.untrack();
    }
});
```

---

### ungroup

Ungroups any grouped shapes in the specified shape group.

#### Signature

**Returns:** `Word.ShapeCollection`

#### Examples

**Example**: Ungroup all shapes in the first shape group found in the document

```typescript
await Word.run(async (context) => {
    // Get the first shape group in the document
    const shapeGroups = context.document.body.shapeGroups;
    shapeGroups.load("items");
    
    await context.sync();
    
    if (shapeGroups.items.length > 0) {
        const firstShapeGroup = shapeGroups.items[0];
        
        // Ungroup the shapes in the shape group
        firstShapeGroup.ungroup();
        
        await context.sync();
        console.log("Shape group has been ungrouped.");
    } else {
        console.log("No shape groups found in the document.");
    }
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

#### Signature

**Returns:** `Word.ShapeGroup`

#### Examples

**Example**: Process a shape group to get its properties, then untrack it to free memory after you're done working with it.

```typescript
await Word.run(async (context) => {
    // Get the first shape group in the document
    const shapeGroup = context.document.body.shapeGroups.getFirst();
    
    // Track the object to work with it
    shapeGroup.track();
    
    // Load properties you need
    shapeGroup.load("id,name");
    await context.sync();
    
    // Use the shape group data
    console.log(`Shape Group: ${shapeGroup.name} (ID: ${shapeGroup.id})`);
    
    // Untrack to release memory after you're done
    shapeGroup.untrack();
    await context.sync();
});
```

---

## Source

- https://docs.microsoft.com/en-us/javascript/api/word
