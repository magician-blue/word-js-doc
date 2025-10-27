# Word.TextFrame

**Package:** `word`

**API Set:** WordApiDesktop 1.2

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents the text frame of a shape object.

## Properties

### autoSizeSetting

**Type:** `None`

The automatic sizing settings for the text frame. A text frame can be set to automatically fit the text to the text frame, to automatically fit the text frame to the text, or not perform any automatic sizing.

#### Examples

**Example**: Configure a shape's text frame to automatically resize to fit its text content

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    const shape = shapes.items[0];
    const textFrame = shape.textFrame;
    
    // Set the text frame to automatically fit the frame to the text
    textFrame.autoSizeSetting = Word.ShapeAutoSize.autoSizeFitToText;
    
    await context.sync();
});
```

---

### bottomMargin

**Type:** `None`

Represents the bottom margin, in points, of the text frame.

#### Examples

**Example**: Set the bottom margin of a shape's text frame to 20 points

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    const shape = shapes.items[0];
    const textFrame = shape.textFrame;
    textFrame.bottomMargin = 20;
    
    await context.sync();
});
```

---

### context

**Type:** `None`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the text frame's request context to verify the connection between the add-in and Word before performing operations on a shape's text frame.

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const textFrame = shape.textFrame;
        
        // Access the context property to verify the connection
        const requestContext = textFrame.context;
        
        // Use the context to perform operations
        textFrame.load("hasText");
        await requestContext.sync();
        
        console.log("Text frame context is connected:", requestContext !== null);
        console.log("Text frame has text:", textFrame.hasText);
    }
});
```

---

### hasText

**Type:** `None`

Specifies if the text frame contains text.

#### Examples

**Example**: Check if a shape's text frame contains text and display an alert with the result

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const textFrame = shape.textFrame;
        textFrame.load("hasText");
        await context.sync();
        
        if (textFrame.hasText) {
            console.log("The text frame contains text");
        } else {
            console.log("The text frame is empty");
        }
    }
});
```

---

### leftMargin

**Type:** `None`

Represents the left margin, in points, of the text frame.

#### Examples

**Example**: Set the left margin of a shape's text frame to 20 points

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    const shape = shapes.items[0];
    const textFrame = shape.textFrame;
    textFrame.leftMargin = 20;
    
    await context.sync();
});
```

---

### noTextRotation

**Type:** `None`

Returns True if text in the text frame shouldn't rotate when the shape is rotated.

#### Examples

**Example**: Prevent text from rotating when a shape is rotated by setting the noTextRotation property to true

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    const shape = shapes.items[0];
    const textFrame = shape.textFrame;
    
    // Set noTextRotation to true so text stays upright when shape rotates
    textFrame.noTextRotation = true;
    
    await context.sync();
    console.log("Text rotation disabled for the shape");
});
```

---

### orientation

**Type:** `None`

Represents the angle to which the text is oriented for the text frame. See Word.ShapeTextOrientation for details.

#### Examples

**Example**: Set the text orientation of a shape's text frame to vertical (rotated 90 degrees)

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    // Get the first shape
    const shape = shapes.items[0];
    const textFrame = shape.textFrame;
    
    // Set the text orientation to vertical (rotated 90 degrees)
    textFrame.orientation = Word.ShapeTextOrientation.rotate90;
    
    await context.sync();
});
```

---

### rightMargin

**Type:** `None`

Represents the right margin, in points, of the text frame.

#### Examples

**Example**: Set the right margin of a shape's text frame to 20 points

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    const shape = shapes.items[0];
    const textFrame = shape.textFrame;
    textFrame.rightMargin = 20;
    
    await context.sync();
});
```

---

### topMargin

**Type:** `None`

Represents the top margin, in points, of the text frame.

#### Examples

**Example**: Set the top margin of a shape's text frame to 20 points

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    const shape = shapes.items[0];
    const textFrame = shape.textFrame;
    textFrame.topMargin = 20;
    
    await context.sync();
});
```

---

### verticalAlignment

**Type:** `None`

Represents the vertical alignment of the text frame. See Word.ShapeTextVerticalAlignment for details.

#### Examples

**Example**: Set the vertical alignment of a shape's text frame to center

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    const shape = shapes.items[0];
    const textFrame = shape.textFrame;
    
    // Set vertical alignment to center
    textFrame.verticalAlignment = Word.ShapeTextVerticalAlignment.center;
    
    await context.sync();
});
```

---

### wordWrap

**Type:** `None`

Determines whether lines break automatically to fit text inside the shape.

#### Examples

**Example**: Disable automatic line wrapping for a shape's text frame so that text extends beyond the shape boundaries instead of breaking into multiple lines

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    // Get the first shape
    const shape = shapes.items[0];
    const textFrame = shape.textFrame;
    
    // Disable word wrap
    textFrame.wordWrap = false;
    
    await context.sync();
});
```

---

## Methods

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `None` (required)

  **Returns:** `None`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `None` (required)

  **Returns:** `None`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `None` (required)

  **Returns:** `None`

#### Examples

**Example**: Load and display the text frame properties of the first shape in the document, including its margins and text orientation.

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    const shape = shapes.getFirst();
    const textFrame = shape.textFrame;
    
    // Load specific properties of the text frame
    textFrame.load("marginTop, marginBottom, marginLeft, marginRight, textOrientation");
    
    // Sync to read the loaded properties
    await context.sync();
    
    // Display the text frame properties
    console.log("Text Frame Properties:");
    console.log(`Top Margin: ${textFrame.marginTop}`);
    console.log(`Bottom Margin: ${textFrame.marginBottom}`);
    console.log(`Left Margin: ${textFrame.marginLeft}`);
    console.log(`Right Margin: ${textFrame.marginRight}`);
    console.log(`Text Orientation: ${textFrame.textOrientation}`);
});
```

---

### set

**Kind:** `write`

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `None` (required)
  - `options`: `None` (required)

  **Returns:** `None`

**Overload 2:**

  **Parameters:**
  - `properties`: `None` (required)

  **Returns:** `None`

#### Examples

**Example**: Configure multiple text frame properties at once to set margins and text orientation for a shape's text frame

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    const shape = shapes.items[0];
    const textFrame = shape.textFrame;
    
    // Set multiple text frame properties at once
    textFrame.set({
        leftMargin: 10,
        rightMargin: 10,
        topMargin: 5,
        bottomMargin: 5,
        verticalAlignment: Word.ShapeTextVerticalAlignment.middle
    });
    
    await context.sync();
    console.log("Text frame properties updated successfully");
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). JSON.stringify, in turn, calls the toJSON method of the object that's passed to it. Whereas the original Word.TextFrame object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.TextFrameData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Serialize a shape's text frame properties to JSON format for logging or data export purposes.

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const textFrame = shape.textFrame;
        
        // Load properties you want to serialize
        textFrame.load("hasText,marginBottom,marginLeft,marginRight,marginTop");
        await context.sync();
        
        // Convert to plain JavaScript object
        const textFrameData = textFrame.toJSON();
        
        // Now you can use the plain object (e.g., log it, send to server, etc.)
        console.log("Text Frame Data:", JSON.stringify(textFrameData, null, 2));
    }
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Track a shape's text frame object across multiple sync calls to safely modify its properties without getting "InvalidObjectPath" errors

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.inlinePictures;
    context.load(shapes);
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const textFrame = shape.textFrame;
        
        // Track the text frame to use it across multiple sync calls
        textFrame.track();
        
        context.load(textFrame, "hasText");
        await context.sync();
        
        // Now we can safely use the textFrame after sync
        if (textFrame.hasText) {
            const textRange = textFrame.textRange;
            context.load(textRange, "text");
            await context.sync();
            
            console.log("Text frame content:", textRange.text);
        }
        
        // Untrack when done
        textFrame.untrack();
    }
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for contex

#### Signature

**Returns:** `None`

#### Examples

**Example**: Load a shape's text frame properties, use them, then release the memory by untracking the object to optimize performance

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    const shape = shapes.items[0];
    const textFrame = shape.textFrame;
    
    // Track the text frame object for changes
    textFrame.track();
    textFrame.load("hasText");
    await context.sync();
    
    // Use the text frame properties
    console.log("Shape has text: " + textFrame.hasText);
    
    // Release memory associated with the tracked object
    textFrame.untrack();
    await context.sync();
});
```

---

## Source

- https://learn.microsoft.com/en-us/javascript/api/word
- https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientobject
- https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets
