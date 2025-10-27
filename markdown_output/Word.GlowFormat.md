# GlowFormat

**Package:** `word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents the glow formatting for the font used by the range of text.

## Properties

### color

**Type:** `Word.ColorFormat`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns a ColorFormat object that represents the color for a glow effect.

#### Examples

**Example**: Set the glow color of selected text to red

```typescript
await Word.run(async (context) => {
    const range = context.document.getSelection();
    const glowFormat = range.font.glow;
    glowFormat.color.set("#FF0000");
    
    await context.sync();
});
```

---

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a GlowFormat object to verify the connection between the add-in and Word before applying glow formatting to selected text.

```typescript
await Word.run(async (context) => {
    const range = context.document.getSelection();
    const glowFormat = range.font.glowFormat;
    
    // Access the request context associated with the GlowFormat object
    const glowContext = glowFormat.context;
    
    // Verify the context is valid by checking if it matches the Word context
    if (glowContext) {
        // Load properties to ensure the context is active
        glowFormat.load("color");
        await context.sync();
        
        console.log("GlowFormat context is connected to Word application");
    }
});
```

---

### radius

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the length of the radius for a glow effect.

#### Examples

**Example**: Set the glow radius to 10 points for the selected text to create a subtle glow effect around the font.

```typescript
await Word.run(async (context) => {
    const range = context.document.getSelection();
    const glowFormat = range.font.glowFormat;
    
    glowFormat.radius = 10;
    
    await context.sync();
});
```

---

### transparency

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the degree of transparency for the glow effect as a value between 0.0 (opaque) and 1.0 (clear).

#### Examples

**Example**: Set the glow effect transparency to 50% (semi-transparent) for the selected text's font

```typescript
await Word.run(async (context) => {
    const range = context.document.getSelection();
    const glowFormat = range.font.glowFormat;
    
    glowFormat.transparency = 0.5;
    
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
  - `options`: `Word.Interfaces.GlowFormatLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.GlowFormat`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.GlowFormat`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.GlowFormat`

#### Examples

**Example**: Get and display the glow color of the selected text's font formatting

```typescript
await Word.run(async (context) => {
    const selection = context.document.getSelection();
    const font = selection.font;
    const glowFormat = font.glowFormat;
    
    // Load the glow format properties
    glowFormat.load("color, radius, transparency");
    
    await context.sync();
    
    console.log("Glow color:", glowFormat.color);
    console.log("Glow radius:", glowFormat.radius);
    console.log("Glow transparency:", glowFormat.transparency);
});
```

---

### set

**Kind:** `configure`

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `Interfaces.GlowFormatUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.GlowFormat` (required)

  **Returns:** `void`

#### Examples

**Example**: Apply glow formatting to selected text by setting multiple glow properties (color and radius) at once

```typescript
await Word.run(async (context) => {
    const selection = context.document.getSelection();
    const glowFormat = selection.font.glowFormat;
    
    glowFormat.set({
        color: "blue",
        radius: 10
    });
    
    await context.sync();
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.GlowFormat object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.GlowFormatData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.GlowFormatData`

#### Examples

**Example**: Get the glow formatting properties of selected text as a plain JavaScript object for logging or data transfer purposes.

```typescript
await Word.run(async (context) => {
    // Get the selected range
    const range = context.document.getSelection();
    
    // Get the glow format of the font
    const glowFormat = range.font.glow;
    
    // Load the glow format properties
    glowFormat.load("color,radius,transparency");
    
    await context.sync();
    
    // Convert to plain JavaScript object
    const glowData = glowFormat.toJSON();
    
    // Log the glow properties as a plain object
    console.log("Glow Format Data:", glowData);
    console.log("Color:", glowData.color);
    console.log("Radius:", glowData.radius);
    console.log("Transparency:", glowData.transparency);
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.GlowFormat`

#### Examples

**Example**: Apply glow formatting to selected text and track the glow format object to maintain reference across multiple sync calls for later modification.

```typescript
await Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.font.load("text");
    
    await context.sync();
    
    // Apply glow formatting
    const glowFormat = selection.font.glow;
    glowFormat.radius = 10;
    glowFormat.color = "blue";
    glowFormat.transparency = 0.5;
    
    // Track the glow format object for use across sync calls
    glowFormat.track();
    
    await context.sync();
    
    // Now we can safely modify the tracked object after sync
    glowFormat.radius = 15;
    glowFormat.color = "red";
    
    await context.sync();
    
    // Untrack when done
    glowFormat.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.GlowFormat`

#### Examples

**Example**: Apply glow formatting to selected text, then untrack the glow format object to free memory after the operation is complete.

```typescript
await Word.run(async (context) => {
    const range = context.document.getSelection();
    const font = range.font;
    const glowFormat = font.glowFormat;
    
    // Track the object to work with it
    glowFormat.track();
    
    // Configure glow properties
    glowFormat.radius = 10;
    glowFormat.color = "blue";
    glowFormat.transparency = 0.5;
    
    await context.sync();
    
    // Release the memory associated with the tracked object
    glowFormat.untrack();
    
    await context.sync();
});
```

---

## Source

- /en-us/javascript/api/word
- /en-us/javascript/api/office/officeextension.clientobject
- /en-us/javascript/api/requirement-sets/word/word-api-requirement-sets
- /en-us/javascript/api/word/word.colorformat
- /en-us/javascript/api/word/word.requestcontext
- /en-us/javascript/api/word/word.interfaces.glowformatloadoptions
- /en-us/javascript/api/word/word.glowformat
- /en-us/javascript/api/office/officeextension.updateoptions
- /en-us/javascript/api/word/word.interfaces.glowformatupdatedata
- /en-us/javascript/api/word/word.interfaces.glowformatdata
- /en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member
