# Word.Shape

**Package:** `word`

**API Set:** WordApiDesktop 1.2

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents a shape in the header, footer, or document body. Currently, only the following shapes are supported: text boxes, geometric shapes, groups, pictures, and canvases.

## Class Examples

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

## Properties

### allowOverlap

**Type:** `boolean`

**Since:** WordApiDesktop 1.2

Specifies whether a given shape can overlap other shapes.

#### Examples

**Example**: Prevent a text box shape from overlapping with other shapes in the document

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document body
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        
        // Set the shape to not allow overlap with other shapes
        shape.allowOverlap = false;
        
        await context.sync();
        console.log("Shape overlap has been disabled");
    }
});
```

---

### altTextDescription

**Type:** `string`

**Since:** WordApiDesktop 1.2

Specifies a string that represents the alternative text associated with the shape.

#### Examples

**Example**: Set the alternative text description for a shape to provide accessibility information for screen readers

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document body
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        
        // Set the alternative text description for the shape
        shape.altTextDescription = "A blue rectangle containing quarterly sales data for the marketing team";
        
        await context.sync();
        console.log("Alternative text description has been set for the shape");
    }
});
```

---

### body

**Type:** `Word.Body`

**Since:** WordApiDesktop 1.2

Represents the body object of the shape. Only applies to text boxes and geometric shapes.

#### Examples

**Example**: Add formatted text content to a text box shape in the document

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document body
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    // Access the first shape (assuming it's a text box)
    const textBoxShape = shapes.items[0];
    
    // Get the body of the shape and insert text
    const shapeBody = textBoxShape.body;
    shapeBody.insertText("This is text inside the shape", Word.InsertLocation.start);
    
    // Format the text in the shape body
    shapeBody.font.color = "blue";
    shapeBody.font.size = 14;
    
    await context.sync();
});
```

---

### canvas

**Type:** `Word.Canvas`

**Since:** WordApiDesktop 1.2

Gets the canvas associated with the shape. An object with its `isNullObject` property set to `true` will be returned if the shape type isn't "Canvas". For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

#### Examples

**Example**: Check if a shape is a canvas and log its width if it is, otherwise log that the shape is not a canvas.

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();

    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const canvas = shape.canvas;
        canvas.load("isNullObject, width");
        await context.sync();

        if (!canvas.isNullObject) {
            console.log(`Canvas width: ${canvas.width}`);
        } else {
            console.log("The shape is not a canvas.");
        }
    }
});
```

---

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the shape's request context to verify the connection to the Word host application and log its diagnostic information.

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document body
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        
        // Access the request context associated with the shape
        const shapeContext = shape.context;
        
        // Use the context to perform operations
        // For example, verify it matches the main context
        console.log("Shape context is connected:", shapeContext !== null);
        console.log("Contexts match:", shapeContext === context);
        
        // The context can be used for sync operations
        shape.load("name,type");
        await shapeContext.sync();
        
        console.log(`Shape name: ${shape.name}, type: ${shape.type}`);
    }
});
```

---

### fill

**Type:** `Word.ShapeFill`

**Since:** WordApiDesktop 1.2

Returns the fill formatting of the shape.

#### Examples

**Example**: Set the fill color of a shape to blue with 80% transparency

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document body
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        
        // Access the fill property and set its color and transparency
        shape.fill.setSolidColor("#0000FF");
        shape.fill.transparency = 0.8;
        
        await context.sync();
    }
});
```

---

### geometricShapeType

**Type:** `Word.GeometricShapeType`

The geometric shape type of the shape. It will be null if isn't a geometric shape.

#### Examples

**Example**: Check if a shape is a geometric shape and display its type, or insert a message if it's not a geometric shape.

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document body
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        shape.load("geometricShapeType");
        await context.sync();
        
        if (shape.geometricShapeType != null) {
            console.log(`Shape is a geometric shape of type: ${shape.geometricShapeType}`);
        } else {
            console.log("Shape is not a geometric shape (may be a picture, text box, or other type)");
        }
    }
});
```

---

### height

**Type:** `number`

The height, in points, of the shape.

#### Examples

**Example**: Set the height of the first shape in the document to 150 points

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        shape.height = 150;
        await context.sync();
    }
});
```

---

### heightRelative

**Type:** `number`

The percentage of shape height to vertical relative size, see [Word.RelativeSize](/en-us/javascript/api/word/word.relativesize). For an inline or child shape, it can't be set.

#### Examples

**Example**: Set a floating shape's height to 50% relative to the page margin height

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document body
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        
        // Set the shape's height to 50% relative to the page margin
        shape.heightRelative = 50;
        
        await context.sync();
        console.log("Shape height set to 50% relative size");
    }
});
```

---

### id

**Type:** `number`

Gets an integer that represents the shape identifier.

#### Examples

**Example**: Get the shape identifier and display it in the console to track which shape is being processed

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document body
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        shape.load("id");
        await context.sync();
        
        // Display the shape identifier
        console.log(`Shape ID: ${shape.id}`);
    }
});
```

---

### isChild

**Type:** `boolean`

Check whether this shape is a child of a group shape or a canvas shape.

#### Examples

**Example**: Check if a shape is a child of a group or canvas, and display an alert message indicating whether it's a child shape or a top-level shape.

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();

    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        shape.load("isChild");
        await context.sync();

        if (shape.isChild) {
            console.log("This shape is a child of a group or canvas.");
        } else {
            console.log("This shape is a top-level shape.");
        }
    }
});
```

---

### left

**Type:** `number`

The distance, in points, from the left side of the shape to the horizontal relative position, see [Word.RelativeHorizontalPosition](/en-us/javascript/api/word/word.relativehorizontalposition). For an inline shape, it will return 0 and can't be set. For a child shape in a canvas or group, it's relative to the top left corner.

#### Examples

**Example**: Position a text box shape 100 points from the left side of the page

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document body
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        
        // Set the shape's left position to 100 points from the left side
        shape.left = 100;
        
        await context.sync();
        console.log("Shape positioned 100 points from the left");
    }
});
```

---

### leftRelative

**Type:** `number`

The relative left position as a percentage from the left side of the shape to the horizontal relative position, see [Word.RelativeHorizontalPosition](/en-us/javascript/api/word/word.relativehorizontalposition). For an inline or child shape, it will return 0 and can't be set.

#### Examples

**Example**: Set a floating text box shape to be positioned at 25% from the left side relative to its horizontal anchor point.

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document body
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        
        // Set the shape's left position to 25% relative to its horizontal anchor
        shape.leftRelative = 25;
        
        await context.sync();
        console.log("Shape positioned at 25% from the left of its anchor point");
    }
});
```

---

### lockAspectRatio

**Type:** `boolean`

Specifies if the aspect ratio of this shape is locked.

#### Examples

**Example**: Lock the aspect ratio of the first shape in the document to prevent distortion when resizing

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.shapes;
    shapes.load("items");
    
    await context.sync();
    
    if (shapes.items.length > 0) {
        const firstShape = shapes.items[0];
        firstShape.lockAspectRatio = true;
        
        await context.sync();
        console.log("Shape aspect ratio has been locked");
    }
});
```

---

### name

**Type:** `string`

The name of the shape.

#### Examples

**Example**: Get the name of the first shape in the document and display it in the console.

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.shapes;
    shapes.load("items");
    
    await context.sync();
    
    if (shapes.items.length > 0) {
        const firstShape = shapes.items[0];
        firstShape.load("name");
        
        await context.sync();
        
        console.log("Shape name: " + firstShape.name);
    }
});
```

---

### parentCanvas

**Type:** `Word.Shape`

Gets the top-level parent canvas shape of this child shape. It will be null if it isn't a child shape of a canvas.

#### Examples

**Example**: Check if a shape is inside a canvas and if so, apply a border to the parent canvas shape.

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document body
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        
        // Get the parent canvas (if it exists)
        const parentCanvas = shape.parentCanvas;
        parentCanvas.load("name");
        await context.sync();
        
        // Check if the shape is inside a canvas
        if (parentCanvas) {
            console.log(`Shape is inside canvas: ${parentCanvas.name}`);
            
            // Apply a border to the parent canvas
            parentCanvas.lineFormat.color = "blue";
            parentCanvas.lineFormat.weight = 2;
            await context.sync();
        } else {
            console.log("Shape is not inside a canvas");
        }
    }
});
```

---

### parentGroup

**Type:** `Word.Shape`

Gets the top-level parent group shape of this child shape. It will be null if it isn't a child shape of a group.

#### Examples

**Example**: Check if a shape is part of a group and if so, apply a border to the top-level parent group shape.

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();

    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        shape.load("parentGroup");
        await context.sync();

        if (shape.parentGroup) {
            // This shape is part of a group, apply border to the parent group
            shape.parentGroup.lineFormat.color = "red";
            shape.parentGroup.lineFormat.weight = 2;
            console.log("Applied border to parent group shape");
        } else {
            // This shape is not part of a group
            console.log("Shape is not part of a group");
        }
        
        await context.sync();
    }
});
```

---

### relativeHorizontalPosition

**Type:** `Word.RelativeHorizontalPosition`

The relative horizontal position of the shape. For an inline shape, it can't be set. For details, see [Word.RelativeHorizontalPosition](/en-us/javascript/api/word/word.relativehorizontalposition).

#### Examples

**Example**: Set a text box shape's horizontal position to be relative to the page margin

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document body
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        
        // Set the relative horizontal position to page margin
        shape.relativeHorizontalPosition = Word.RelativeHorizontalPosition.page;
        
        await context.sync();
        console.log("Shape horizontal position set relative to page");
    }
});
```

---

### relativeHorizontalSize

**Type:** `Word.RelativeHorizontalSize`

The relative horizontal size of the shape. For an inline or child shape, it can't be set. For details, see [Word.RelativeHorizontalPosition](/en-us/javascript/api/word/word.relativehorizontalposition).

#### Examples

**Example**: Get the relative horizontal size setting of a shape in the document to determine if it's sized relative to the margin, page, or other reference.

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document body
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        shape.load("relativeHorizontalSize");
        await context.sync();
        
        console.log("Relative horizontal size: " + shape.relativeHorizontalSize);
        // Possible values: "margin", "page", "leftMargin", "rightMargin", 
        // "insideMargin", "outsideMargin"
    }
});
```

---

### relativeVerticalPosition

**Type:** `Word.RelativeVerticalPosition`

The relative vertical position of the shape. For an inline shape, it can't be set. For details, see [Word.RelativeVerticalPosition](/en-us/javascript/api/word/word.relativeverticalposition).

#### Examples

**Example**: Get the relative vertical position of a shape in the document and display it in the console.

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document body
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        shape.load("relativeVerticalPosition");
        await context.sync();
        
        console.log("Shape's relative vertical position: " + shape.relativeVerticalPosition);
        // Possible values: "margin", "page", "paragraph", "line", "topMargin", "bottomMargin", "insideMargin", "outsideMargin"
    }
});
```

---

### relativeVerticalSize

**Type:** `Word.RelativeVerticalSize`

The relative vertical size of the shape. For an inline or child shape, it can't be set. For details, see [Word.RelativeHorizontalPosition](/en-us/javascript/api/word/word.relativehorizontalposition).

#### Examples

**Example**: Get the relative vertical size settings of a shape in the document to check its vertical sizing configuration.

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document body
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const relativeVerticalSize = shape.relativeVerticalSize;
        
        // Load properties of the relative vertical size
        relativeVerticalSize.load("relativeHeight, sizeRelativeTo");
        await context.sync();
        
        console.log("Relative Height: " + relativeVerticalSize.relativeHeight);
        console.log("Size Relative To: " + relativeVerticalSize.sizeRelativeTo);
    }
});
```

---

### rotation

**Type:** `number`

Specifies the rotation, in degrees, of the shape. Not applicable to Canvas shape.

#### Examples

**Example**: Rotate a shape in the document by 45 degrees clockwise

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.shapes;
    shapes.load("items");
    
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        shape.rotation = 45;
        
        await context.sync();
    }
});
```

---

### shapeGroup

**Type:** `Word.ShapeGroup`

Gets the shape group associated with the shape. An object with its isNullObject property set to true will be returned if the shape type isn't "GroupShape". For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

#### Examples

**Example**: Check if a shape is a group and log the number of shapes within the group, or indicate if it's not a group shape.

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document body
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const shapeGroup = shape.shapeGroup;
        shapeGroup.load("isNullObject, shapes");
        await context.sync();
        
        if (shapeGroup.isNullObject) {
            console.log("This shape is not a group shape.");
        } else {
            shapeGroup.shapes.load("items");
            await context.sync();
            console.log(`This is a group shape containing ${shapeGroup.shapes.items.length} shapes.`);
        }
    }
});
```

---

### textFrame

**Type:** `Word.TextFrame`

Gets the text frame object of the shape.

#### Examples

**Example**: Get the text frame of a shape and set its text content to "Hello World"

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document body
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        
        // Get the text frame of the shape
        const textFrame = shape.textFrame;
        textFrame.load("textRange");
        await context.sync();
        
        // Set text content in the text frame
        textFrame.textRange.text = "Hello World";
        await context.sync();
    }
});
```

---

### textWrap

**Type:** `Word.TextWrap`

Returns the text wrap formatting of the shape.

#### Examples

**Example**: Get the text wrap type of the first shape in the document and display it to the user.

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document body
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const textWrap = shape.textWrap;
        textWrap.load("type");
        await context.sync();
        
        console.log(`Shape text wrap type: ${textWrap.type}`);
        // Possible values: "inline", "square", "tight", "through", "topBottom", "behind", "inFrontOf"
    }
});
```

---

### top

**Type:** `number`

The distance, in points, from the top edge of the shape to the vertical relative position (see [Word.RelativeVerticalPosition](/en-us/javascript/api/word/word.relativeverticalposition)). For an inline shape, it will return 0 and can't be set. For a child shape in a canvas or group, it's relative to the top left corner.

#### Examples

**Example**: Get the distance from the top edge of a floating shape to its vertical reference point and display it in the console.

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document body
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        shape.load("top, name");
        await context.sync();
        
        console.log(`Shape "${shape.name}" is ${shape.top} points from the top edge`);
    }
});
```

---

### topRelative

**Type:** `number`

The relative top position as a percentage from the top edge of the shape to the vertical relative position, see [Word.RelativeVerticalPosition](/en-us/javascript/api/word/word.relativeverticalposition). For an inline or child shape, it will return 0 and can't be set.

#### Examples

**Example**: Set a floating shape's vertical position to 25% from the top of the page using relative positioning

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document body
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        
        // Set the shape's vertical position to 25% from the top
        shape.topRelative = 25;
        
        await context.sync();
        console.log("Shape positioned at 25% from the top");
    }
});
```

---

### type

**Type:** `Word.ShapeType`

Gets the shape type. Currently, only the following shapes are supported: text boxes, geometric shapes, groups, pictures, and canvases.

#### Examples

**Example**: Get the type of the first shape in the document and display it to the user

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.shapes;
    shapes.load("items");
    
    await context.sync();
    
    if (shapes.items.length > 0) {
        const firstShape = shapes.items[0];
        firstShape.load("type");
        
        await context.sync();
        
        console.log(`Shape type: ${firstShape.type}`);
        // Type will be one of: "TextBox", "GeometricShape", "Group", "Picture", or "Canvas"
    }
});
```

---

### visible

**Type:** `boolean`

Specifies if the shape is visible. Not applicable to inline shapes.

#### Examples

**Example**: Hide all floating shapes in the document body that are currently visible

```typescript
await Word.run(async (context) => {
    // Get all shapes in the document body
    const shapes = context.document.body.shapes;
    shapes.load("items");
    
    await context.sync();
    
    // Hide all visible shapes
    for (let i = 0; i < shapes.items.length; i++) {
        const shape = shapes.items[i];
        shape.load("visible");
        await context.sync();
        
        if (shape.visible) {
            shape.visible = false;
        }
    }
    
    await context.sync();
});
```

---

### width

**Type:** `number`

The width, in points, of the shape.

#### Examples

**Example**: Set the width of the first shape in the document body to 200 points

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.shapes;
    shapes.load("items");
    
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        shape.width = 200;
        
        await context.sync();
    }
});
```

---

### widthRelative

**Type:** `number`

The percentage of shape width to horizontal relative size, see [Word.RelativeSize](/en-us/javascript/api/word/word.relativesize). For an inline or child shape, it can't be set.

#### Examples

**Example**: Set a floating shape's width to 50% of the page width using relative sizing

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document body
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        
        // Set the shape's width to 50% of the page width
        shape.widthRelative = 50;
        
        await context.sync();
        console.log("Shape width set to 50% of page width");
    }
});
```

---

## Methods

### delete

**Kind:** `delete`

Deletes the shape and its content.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Delete all shapes in the document body that contain the text "Draft" in their name

```typescript
await Word.run(async (context) => {
    // Get all shapes in the document body
    const shapes = context.document.body.shapes;
    shapes.load("items");
    
    await context.sync();
    
    // Find and delete shapes with "Draft" in their name
    for (let i = 0; i < shapes.items.length; i++) {
        const shape = shapes.items[i];
        shape.load("name");
        await context.sync();
        
        if (shape.name.includes("Draft")) {
            shape.delete();
        }
    }
    
    await context.sync();
    console.log("Shapes containing 'Draft' have been deleted.");
});
```

---

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `string[] | OfficeExtension.LoadOption` (required)

  **Returns:** `None`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string[]` (required)

  **Returns:** `None`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ [propertyName: string]: string }` (required)

  **Returns:** `None`

#### Examples

**Example**: Load and display the width and height properties of the first shape in the document body.

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document body
    const shapes = context.document.body.shapes;
    const firstShape = shapes.getFirst();
    
    // Load the width and height properties
    firstShape.load("width, height");
    
    // Sync to execute the load command
    await context.sync();
    
    // Display the loaded properties
    console.log(`Shape width: ${firstShape.width}`);
    console.log(`Shape height: ${firstShape.height}`);
});
```

---

### moveHorizontally

Moves the shape horizontally by the number of points.

#### Signature

**Parameters:**
- `distance`: `number` (required)

**Returns:** `None`

#### Examples

**Example**: Move a text box shape 50 points to the right in the document

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document body
    const shapes = context.document.body.shapes;
    shapes.load("items");
    
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        
        // Move the shape 50 points to the right
        shape.moveHorizontally(50);
        
        await context.sync();
    }
});
```

---

### moveVertically

Moves the shape vertically by the number of points.

#### Signature

**Parameters:**
- `distance`: `number` (required)

**Returns:** `None`

#### Examples

**Example**: Move a text box shape down by 50 points from its current position

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document body
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        
        // Move the shape down by 50 points
        shape.moveVertically(50);
        
        await context.sync();
    }
});
```

---

### scaleHeight

Scales the height of the shape by a specified factor. For images, you can indicate whether you want to scale the shape relative to the original or the current size. Shapes other than pictures are always scaled relative to their current height.

#### Signature

**Parameters:**
- `scaleFactor`: `number` (required)
- `scaleType`: `Word.ShapeScaleType` (required)
- `scaleFrom`: `Word.ShapeScaleFrom` (required)

**Returns:** `None`

#### Examples

**Example**: Scale a text box shape in the document to 150% of its current height, scaling from the top edge

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document body
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        
        // Scale the shape height to 150% of current size, from the top
        shape.scaleHeight(
            1.5,                                    // scaleFactor: 150% of current height
            Word.ScaleType.current,                 // scaleType: relative to current size
            Word.ScaleFrom.topLeft                  // scaleFrom: scale from top edge
        );
        
        await context.sync();
        console.log("Shape height scaled successfully");
    }
});
```

---

### scaleWidth

Scales the width of the shape by a specified factor. For images, you can indicate whether you want to scale the shape relative to the original or the current size. Shapes other than pictures are always scaled relative to their current height.

#### Signature

**Parameters:**
- `scaleFactor`: `number` (required)
- `scaleType`: `Word.ShapeScaleType` (required)
- `scaleFrom`: `Word.ShapeScaleFrom` (required)

**Returns:** `None`

#### Examples

**Example**: Scale a text box shape's width to 150% of its current size, anchoring the scaling from the center of the shape

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document body
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        
        // Scale the width to 150% of current size, from the center
        shape.scaleWidth(
            1.5,                                    // scaleFactor: 150% of current width
            Word.ShapeScaleType.currentSize,        // scaleType: relative to current size
            Word.ShapeScaleFrom.scaleFromMiddle     // scaleFrom: anchor from center
        );
        
        await context.sync();
        console.log("Shape width scaled successfully");
    }
});
```

---

### select

Selects the shape.

#### Signature

**Parameters:**
- `selectMultipleShapes`: `boolean` (required)

**Returns:** `None`

#### Examples

**Example**: Select a specific shape in the document to prepare it for formatting or manipulation

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document body
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        
        // Select the shape
        shape.select();
        
        await context.sync();
        console.log("Shape selected successfully");
    }
});
```

---

### set

**Kind:** `configure`

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `Word.Interfaces.ShapeUpdateData` (required)
  - `options`: `string[]` (required)

  **Returns:** `None`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.Shape` (required)

  **Returns:** `None`

#### Examples

**Example**: Update multiple properties of a shape at once by setting its fill color, line color, and dimensions

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document body
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    const shape = shapes.items[0];
    
    // Set multiple properties at once
    shape.set({
        fill: {
            type: Word.ShapeFillType.solid,
            color: "#4472C4"
        },
        lineFormat: {
            color: "#FF0000",
            weight: 2
        },
        width: 200,
        height: 150
    });
    
    await context.sync();
    console.log("Shape properties updated successfully");
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.Shape object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.ShapeData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.ShapeData`

#### Examples

**Example**: Get a shape from the document, load its properties, and serialize it to JSON format for logging or data transfer purposes.

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document body
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        
        // Load properties you want to serialize
        shape.load("id,name,type,width,height,left,top");
        await context.sync();
        
        // Convert the shape to a plain JavaScript object
        const shapeData = shape.toJSON();
        
        // Now you can use the plain object (e.g., log it, send it to a server)
        console.log("Shape data:", JSON.stringify(shapeData, null, 2));
    }
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Track a text box shape across multiple sync calls to modify its properties without getting an "InvalidObjectPath" error

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document body
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    const shape = shapes.items[0];
    
    // Track the shape to use it across multiple sync calls
    shape.track();
    
    // Load properties
    shape.load("type,name");
    await context.sync();
    
    console.log(`Shape type: ${shape.type}, Name: ${shape.name}`);
    
    // Modify shape properties in a subsequent sync call
    // Without track(), this would throw an "InvalidObjectPath" error
    if (shape.type === Word.ShapeType.textBox) {
        const textFrame = shape.textFrame;
        textFrame.load("textRange");
        await context.sync();
        
        textFrame.textRange.text = "Updated text content";
        await context.sync();
    }
    
    // Untrack when done to free up memory
    shape.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Process multiple shapes in a document, track them for batch operations, then untrack them to free memory after processing is complete.

```typescript
await Word.run(async (context) => {
    // Get all shapes in the document body
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();

    // Track shapes for processing
    const trackedShapes: Word.Shape[] = [];
    for (let i = 0; i < shapes.items.length; i++) {
        const shape = shapes.items[i];
        shape.track();
        trackedShapes.push(shape);
        shape.load("name,type");
    }
    await context.sync();

    // Process the shapes (e.g., log their properties)
    trackedShapes.forEach(shape => {
        console.log(`Shape: ${shape.name}, Type: ${shape.type}`);
    });

    // Untrack all shapes to release memory
    trackedShapes.forEach(shape => {
        shape.untrack();
    });
    await context.sync();
});
```

---

## Source

- /en-us/javascript/api/word
