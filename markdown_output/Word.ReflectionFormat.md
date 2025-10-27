# Word.ReflectionFormat

**Package:** `word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents the reflection formatting for a shape in Word.

## Properties

### blur

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the degree of blur effect applied to the ReflectionFormat object as a value between 0.0 and 100.0.

#### Examples

**Example**: Set the reflection blur effect to 50 points for a selected shape to create a moderately blurred reflection

```typescript
await Word.run(async (context) => {
    const shapes = context.document.getSelection().inlinePictures;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const reflection = shape.reflection;
        reflection.blur = 50;
        
        await context.sync();
    }
});
```

---

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a shape's reflection format to verify the connection to the Word host application before applying reflection settings.

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.getShapes();
    const shape = shapes.getFirst();
    
    // Load the reflection format
    shape.load("reflection");
    await context.sync();
    
    // Access the request context from the reflection format
    const reflectionContext = shape.reflection.context;
    
    // Verify the context is connected to the same Word application
    if (reflectionContext === context) {
        console.log("Reflection format context is properly connected to Word");
        
        // Now safe to modify reflection properties
        shape.reflection.transparency = 0.5;
        shape.reflection.blur = 10;
        await context.sync();
    }
});
```

---

### offset

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the amount of separation, in points, of the reflected image from the shape.

#### Examples

**Example**: Set the reflection offset to 5 points to create separation between a shape and its reflected image

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.inlinePictures;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const reflection = shape.imageFormat.reflection;
        reflection.offset = 5;
        
        await context.sync();
    }
});
```

---

### size

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the size of the reflection as a percentage of the reflected shape from 0 to 100.

#### Examples

**Example**: Set the reflection size of a shape to 75% of the original shape's size

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    const shape = shapes.items[0];
    shape.reflection.size = 75;
    
    await context.sync();
});
```

---

### transparency

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the degree of transparency for the reflection effect as a value between 0.0 (opaque) and 1.0 (clear).

#### Examples

**Example**: Set the reflection transparency of a shape to 50% (semi-transparent) to create a subtle reflection effect

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.inlinePictures;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const reflection = shape.reflection;
        reflection.transparency = 0.5;
        
        await context.sync();
    }
});
```

---

### type

**Type:** `Word.ReflectionType | "Mixed" | "None" | "Type1" | "Type2" | "Type3" | "Type4" | "Type5" | "Type6" | "Type7" | "Type8" | "Type9"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies a ReflectionType value that represents the type and direction of the lighting for a shape reflection.

#### Examples

**Example**: Set the reflection type of a selected shape to "Type4" to apply a medium-intensity reflection effect.

```typescript
await Word.run(async (context) => {
    const shapes = context.document.getSelection().inlinePictures;
    shapes.load("items");
    await context.sync();

    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const reflectionFormat = shape.imageFormat.reflection;
        reflectionFormat.type = Word.ReflectionType.type4;
        
        await context.sync();
        console.log("Reflection type set to Type4");
    }
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
  - `options`: `Word.Interfaces.ReflectionFormatLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.ReflectionFormat`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.ReflectionFormat`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.ReflectionFormat`

#### Examples

**Example**: Load and read the transparency property of a shape's reflection formatting to check if the reflection is visible.

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    const shape = shapes.items[0];
    const reflectionFormat = shape.reflection;
    
    // Load the reflection format properties
    reflectionFormat.load("transparency");
    await context.sync();
    
    // Read the loaded property
    console.log(`Reflection transparency: ${reflectionFormat.transparency}`);
});
```

---

### set

**Kind:** `write`

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `Interfaces.ReflectionFormatUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.ReflectionFormat` (required)

  **Returns:** `void`

#### Examples

**Example**: Apply reflection formatting to a shape by setting multiple reflection properties at once, including blur, distance, size, and transparency.

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.getShapes();
    const shape = shapes.getFirst();
    
    // Set multiple reflection properties at once
    shape.reflection.set({
        blur: 5,
        distance: 3,
        size: 80,
        transparency: 0.5,
        type: Word.ReflectionType.tight
    });
    
    await context.sync();
    console.log("Reflection formatting applied to shape");
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.ReflectionFormat object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.ReflectionFormatData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.ReflectionFormatData`

#### Examples

**Example**: Get the reflection formatting properties of a shape as a plain JavaScript object and log it to the console for inspection or serialization.

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const reflection = shape.reflection;
        
        // Load reflection properties
        reflection.load("transparency,size,type,blur");
        await context.sync();
        
        // Convert to plain JavaScript object
        const reflectionData = reflection.toJSON();
        
        // Log the plain object (useful for debugging or serialization)
        console.log("Reflection properties:", reflectionData);
        console.log("Transparency:", reflectionData.transparency);
        console.log("Size:", reflectionData.size);
    }
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.ReflectionFormat`

#### Examples

**Example**: Apply reflection formatting to a shape and track it across multiple sync calls to maintain the reference while modifying its properties

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.inlinePictures;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const reflection = shape.reflection;
        
        // Track the reflection object to use it across multiple sync calls
        reflection.track();
        
        // First sync: load current properties
        reflection.load("transparency");
        await context.sync();
        
        console.log("Current transparency: " + reflection.transparency);
        
        // Second sync: modify properties
        reflection.transparency = 0.5;
        await context.sync();
        
        // Untrack when done
        reflection.untrack();
    }
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.ReflectionFormat`

#### Examples

**Example**: Apply reflection formatting to a shape, then untrack the reflection format object to free memory after the formatting is complete.

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        
        // Get and configure the reflection format
        const reflectionFormat = shape.reflection;
        context.trackedObjects.add(reflectionFormat);
        
        reflectionFormat.transparency = 0.5;
        reflectionFormat.size = 75;
        reflectionFormat.blur = 10;
        
        await context.sync();
        
        // Release memory after we're done with the reflection format
        reflectionFormat.untrack();
        await context.sync();
        
        console.log("Reflection formatting applied and object untracked");
    }
});
```

---

## Source

- https://learn.microsoft.com/en-us/javascript/api/word/word.reflectionformat
