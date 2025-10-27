# OleFormat

**Package:** `word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents the OLE characteristics (other than linking) for an OLE object, ActiveX control, or field.

## Properties

### classType

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the class type for the specified OLE object, picture, or field.

#### Examples

**Example**: Get the class type of the first OLE object in the document to identify what type of embedded object it is (e.g., Excel worksheet, PowerPoint slide, etc.)

```typescript
await Word.run(async (context) => {
    // Get the first inline shape (which may be an OLE object)
    const inlineShapes = context.document.body.inlineShapes;
    inlineShapes.load("items");
    
    await context.sync();
    
    if (inlineShapes.items.length > 0) {
        const firstShape = inlineShapes.items[0];
        const oleFormat = firstShape.oleFormat;
        oleFormat.load("classType");
        
        await context.sync();
        
        console.log("OLE Object Class Type: " + oleFormat.classType);
        // Example output: "Excel.Sheet.12" or "PowerPoint.Show.12"
    }
});
```

---

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the OleFormat's request context to load and read properties of an OLE object in the document

```typescript
await Word.run(async (context) => {
    // Get the first OLE object in the document
    const oleObjects = context.document.body.inlinePictures;
    oleObjects.load("items");
    await context.sync();
    
    if (oleObjects.items.length > 0) {
        const firstOleObject = oleObjects.items[0];
        const oleFormat = firstOleObject.oleFormat;
        
        // Access the request context from the OleFormat object
        const oleContext = oleFormat.context;
        
        // Use the context to load OLE format properties
        oleFormat.load("iconName,iconIndex");
        await oleContext.sync();
        
        console.log(`OLE Icon Name: ${oleFormat.iconName}`);
        console.log(`OLE Icon Index: ${oleFormat.iconIndex}`);
    }
});
```

---

### iconIndex

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the icon that is used when the `displayAsIcon` property is `true`.

#### Examples

**Example**: Set an embedded Excel worksheet to display as an icon and specify which icon to use (icon at index 2 from the source application's icon list)

```typescript
await Word.run(async (context) => {
    // Get the first inline shape (assuming it's an OLE object)
    const inlineShape = context.document.body.inlineShapes.getFirst();
    const oleFormat = inlineShape.oleFormat;
    
    // Load the oleFormat properties
    oleFormat.load("displayAsIcon");
    await context.sync();
    
    // Set to display as icon and specify icon index 2
    oleFormat.displayAsIcon = true;
    oleFormat.iconIndex = 2;
    
    await context.sync();
    console.log("OLE object set to display as icon with icon index 2");
});
```

---

### iconLabel

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the text displayed below the icon for the OLE object.

#### Examples

**Example**: Set the icon label text to "Financial Report 2024" for the first OLE object in the document

```typescript
await Word.run(async (context) => {
    // Get the first inline shape (OLE object) in the document
    const inlineShape = context.document.body.inlineShapes.getFirst();
    
    // Access the OLE format
    const oleFormat = inlineShape.oleFormat;
    
    // Set the icon label text
    oleFormat.iconLabel = "Financial Report 2024";
    
    await context.sync();
});
```

---

### iconName

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the program file in which the icon for the OLE object is stored.

#### Examples

**Example**: Get the icon program file name for the first OLE object in the document and display it in the console.

```typescript
await Word.run(async (context) => {
    // Get the first inline shape (OLE object) in the document
    const inlineShapes = context.document.body.inlineShapes;
    inlineShapes.load("items");
    await context.sync();
    
    if (inlineShapes.items.length > 0) {
        const oleObject = inlineShapes.items[0];
        const oleFormat = oleObject.oleFormat;
        oleFormat.load("iconName");
        await context.sync();
        
        console.log("Icon program file: " + oleFormat.iconName);
    } else {
        console.log("No OLE objects found in the document.");
    }
});
```

---

### iconPath

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the path of the file in which the icon for the OLE object is stored.

#### Examples

**Example**: Retrieve and display the icon file path from an OLE object embedded in the document.

```typescript
await Word.run(async (context) => {
    // Get the first inline shape (OLE object) in the document
    const inlineShapes = context.document.body.inlineShapes;
    inlineShapes.load("items");
    await context.sync();
    
    if (inlineShapes.items.length > 0) {
        const oleObject = inlineShapes.items[0];
        const oleFormat = oleObject.oleFormat;
        oleFormat.load("iconPath");
        await context.sync();
        
        // Display the icon path
        console.log("OLE object icon path: " + oleFormat.iconPath);
    } else {
        console.log("No OLE objects found in the document.");
    }
});
```

---

### isDisplayedAsIcon

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets whether the specified object is displayed as an icon.

#### Examples

**Example**: Check if an OLE object in the document is displayed as an icon and log the result to the console

```typescript
await Word.run(async (context) => {
    // Get the first inline shape (OLE object) in the document
    const inlineShapes = context.document.body.inlineShapes;
    inlineShapes.load("items");
    
    await context.sync();
    
    if (inlineShapes.items.length > 0) {
        const firstShape = inlineShapes.items[0];
        const oleFormat = firstShape.oleFormat;
        oleFormat.load("isDisplayedAsIcon");
        
        await context.sync();
        
        console.log(`OLE object is displayed as icon: ${oleFormat.isDisplayedAsIcon}`);
    } else {
        console.log("No OLE objects found in the document");
    }
});
```

---

### isFormattingPreservedOnUpdate

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether formatting done in Microsoft Word to the linked OLE object is preserved.

#### Examples

**Example**: Check if formatting preservation is enabled for a linked OLE object and display the result in the console

```typescript
await Word.run(async (context) => {
    // Get the first inline shape (OLE object) in the document
    const inlineShapes = context.document.body.inlineShapes;
    inlineShapes.load("items");
    
    await context.sync();
    
    if (inlineShapes.items.length > 0) {
        const oleFormat = inlineShapes.items[0].oleFormat;
        oleFormat.load("isFormattingPreservedOnUpdate");
        
        await context.sync();
        
        console.log(`Formatting preserved on update: ${oleFormat.isFormattingPreservedOnUpdate}`);
    } else {
        console.log("No OLE objects found in the document");
    }
});
```

---

### label

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets a string that's used to identify the portion of the source file that's being linked.

#### Examples

**Example**: Get the label that identifies the portion of the source file being linked for an OLE object in the document and display it in the console.

```typescript
await Word.run(async (context) => {
    // Get the first inline shape (OLE object) in the document
    const inlineShapes = context.document.body.inlineShapes;
    inlineShapes.load("items");
    
    await context.sync();
    
    if (inlineShapes.items.length > 0) {
        const oleFormat = inlineShapes.items[0].oleFormat;
        oleFormat.load("label");
        
        await context.sync();
        
        console.log("OLE object label: " + oleFormat.label);
    } else {
        console.log("No OLE objects found in the document.");
    }
});
```

---

### progID

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the programmatic identifier (`ProgId`) for the specified OLE object.

#### Examples

**Example**: Get the programmatic identifier (ProgId) of the first OLE object in the document to determine what type of embedded object it is (e.g., Excel worksheet, PowerPoint slide).

```typescript
await Word.run(async (context) => {
    // Get the first inline shape (which may be an OLE object)
    const inlineShapes = context.document.body.inlineShapes;
    inlineShapes.load("items");
    await context.sync();

    if (inlineShapes.items.length > 0) {
        const firstShape = inlineShapes.items[0];
        const oleFormat = firstShape.oleFormat;
        oleFormat.load("progID");
        await context.sync();

        // Display the ProgId (e.g., "Excel.Sheet.12", "PowerPoint.Slide.12")
        console.log("OLE Object ProgId: " + oleFormat.progID);
    } else {
        console.log("No OLE objects found in the document.");
    }
});
```

---

## Methods

### activate

Activates the `OleFormat` object.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Activate an embedded OLE object (such as an Excel spreadsheet) in the document to open it for editing

```typescript
await Word.run(async (context) => {
    // Get the first inline shape in the document
    const inlineShapes = context.document.body.inlineShapes;
    inlineShapes.load("items");
    await context.sync();
    
    if (inlineShapes.items.length > 0) {
        const firstShape = inlineShapes.items[0];
        
        // Get the OLE format of the shape
        const oleFormat = firstShape.oleFormat;
        oleFormat.load("classType");
        await context.sync();
        
        // Activate the OLE object to open it for editing
        oleFormat.activate();
        await context.sync();
        
        console.log("OLE object activated successfully");
    }
});
```

---

### activateAs

**Kind:** `configure`

Sets the Windows registry value that determines the default application used to activate the specified OLE object.

#### Signature

**Parameters:**
- `classType`: `string` (required)
  The class type to activate as.

**Returns:** `void`

#### Examples

**Example**: Change an embedded Excel worksheet OLE object to activate as a different application (e.g., Notepad) by modifying its class type in the Windows registry.

```typescript
await Word.run(async (context) => {
    // Get the first inline shape (OLE object) in the document
    const inlineShapes = context.document.body.inlineShapes;
    inlineShapes.load("items");
    await context.sync();
    
    if (inlineShapes.items.length > 0) {
        const oleObject = inlineShapes.items[0];
        const oleFormat = oleObject.oleFormat;
        
        // Change the OLE object to activate as Notepad instead of its default application
        oleFormat.activateAs("Notepad.Document");
        
        await context.sync();
        console.log("OLE object activation class type changed successfully");
    }
});
```

---

### doVerb

Requests that the OLE object perform one of its available verbs.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `verbIndex`: `Word.OleVerb` (optional)
    Optional. The index of the verb to perform.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `verbIndex`: `"Primary" | "Show" | "Open" | "Hide" | "UiActivate" | "InPlaceActivate" | "DiscardUndoState"` (optional)
    Optional. The index of the verb to perform.

  **Returns:** `void`

#### Examples

**Example**: Activate an embedded OLE object (such as an Excel spreadsheet) for in-place editing by performing its primary verb.

```typescript
await Word.run(async (context) => {
    // Get the first inline shape in the document (assuming it's an OLE object)
    const firstInlineShape = context.document.body.inlinePictures.getFirst();
    
    // Get the OLE format of the inline shape
    const oleFormat = firstInlineShape.oleFormat;
    
    // Load the OLE format to ensure it exists
    oleFormat.load("classType");
    
    await context.sync();
    
    // Perform the primary verb (0 = primary action, typically "Edit" or "Open")
    oleFormat.doVerb(0);
    
    await context.sync();
});
```

---

### edit

Opens the OLE object for editing in the application it was created in.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Open an embedded Excel spreadsheet OLE object for editing in Microsoft Excel

```typescript
await Word.run(async (context) => {
    // Get the first inline shape in the document (assuming it's an OLE object)
    const inlineShapes = context.document.body.inlineShapes;
    inlineShapes.load("items");
    
    await context.sync();
    
    if (inlineShapes.items.length > 0) {
        const firstShape = inlineShapes.items[0];
        
        // Get the OLE format of the shape
        const oleFormat = firstShape.oleFormat;
        oleFormat.load("classType");
        
        await context.sync();
        
        // Open the OLE object for editing in its native application
        oleFormat.edit();
        
        await context.sync();
    }
});
```

---

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.OleFormatLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.OleFormat`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.OleFormat`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.OleFormat`

#### Examples

**Example**: Load and display the icon name property of the first OLE object in the document

```typescript
await Word.run(async (context) => {
    // Get the first OLE object in the document
    const oleObjects = context.document.body.inlinePictures;
    oleObjects.load("items");
    await context.sync();
    
    // Access the first OLE object's format
    const firstOleObject = oleObjects.items[0];
    const oleFormat = firstOleObject.getOleObjectOrNullObject();
    
    // Load specific properties of the OLE format
    oleFormat.load("iconName");
    await context.sync();
    
    // Display the icon name
    if (!oleFormat.isNullObject) {
        console.log("OLE Object Icon Name: " + oleFormat.iconName);
    } else {
        console.log("No OLE object found");
    }
});
```

---

### open

Opens the `OleFormat` object.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Open an embedded OLE object (such as an Excel spreadsheet) in the first shape of the document for editing.

```typescript
await Word.run(async (context) => {
    // Get the first inline shape in the document
    const firstShape = context.document.body.inlinePictures.getFirst();
    
    // Get the OLE format of the shape
    const oleFormat = firstShape.oleFormat;
    
    // Load the OLE format to check if it exists
    oleFormat.load("iconName");
    
    await context.sync();
    
    // Open the OLE object for editing
    oleFormat.open();
    
    await context.sync();
});
```

---

### set

**Kind:** `write`

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `Interfaces.OleFormatUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.OleFormat` (required)

  **Returns:** `void`

#### Examples

**Example**: Configure multiple OLE object properties at once by setting the icon visibility and display as icon settings for an embedded OLE object in the document.

```typescript
await Word.run(async (context) => {
    // Get the first inline shape (OLE object) in the document
    const inlineShape = context.document.body.inlineShapes.getFirst();
    const oleFormat = inlineShape.oleFormat;
    
    // Set multiple OLE format properties at once
    oleFormat.set({
        displayAsIcon: true,
        iconPath: "C:\\Icons\\custom.ico"
    });
    
    await context.sync();
    console.log("OLE format properties updated successfully");
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.OleFormat` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.OleFormatData`) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.OleFormatData`

#### Examples

**Example**: Serialize an OLE object's format properties to JSON for logging or data export purposes.

```typescript
await Word.run(async (context) => {
    // Get the first inline shape (which could be an OLE object)
    const inlineShapes = context.document.body.inlineShapes;
    inlineShapes.load("items");
    await context.sync();
    
    if (inlineShapes.items.length > 0) {
        const firstShape = inlineShapes.items[0];
        const oleFormat = firstShape.oleFormat;
        
        // Load properties you want to serialize
        oleFormat.load("iconName,iconIndex,displayAsIcon");
        await context.sync();
        
        // Convert to plain JavaScript object
        const oleFormatData = oleFormat.toJSON();
        
        // Now you can use JSON.stringify or log the data
        console.log("OLE Format Data:", JSON.stringify(oleFormatData, null, 2));
        
        // The plain object can be easily stored or transmitted
        return oleFormatData;
    }
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.OleFormat`

#### Examples

**Example**: Track an OLE object (like an embedded Excel chart) across multiple sync calls to maintain its reference while modifying its properties

```typescript
await Word.run(async (context) => {
    // Get the first inline shape (which could be an OLE object)
    const inlineShape = context.document.body.inlineShapes.getFirst();
    inlineShape.load("oleFormat");
    await context.sync();
    
    // Track the OLE format object to use it across multiple sync calls
    const oleFormat = inlineShape.oleFormat;
    oleFormat.track();
    
    // Load properties of the OLE object
    oleFormat.load("iconName,iconIndex");
    await context.sync();
    
    // Now we can safely use the oleFormat object after another sync
    console.log("OLE Icon Name: " + oleFormat.iconName);
    console.log("OLE Icon Index: " + oleFormat.iconIndex);
    
    // Untrack when done to free up memory
    oleFormat.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

#### Signature

**Returns:** `Word.OleFormat`

#### Examples

**Example**: Track an OLE object to work with it across multiple sync operations, then untrack it to release memory when done processing.

```typescript
await Word.run(async (context) => {
    // Get the first inline shape (which could be an OLE object)
    const inlineShape = context.document.body.inlineShapes.getFirst();
    const oleFormat = inlineShape.oleFormat;
    
    // Track the OLE format object for use across multiple syncs
    oleFormat.track();
    
    // Load properties
    oleFormat.load("progId");
    await context.sync();
    
    // Use the OLE format object
    console.log("OLE Program ID: " + oleFormat.progId);
    
    // Do more work with the tracked object...
    await context.sync();
    
    // When done, untrack to release memory
    oleFormat.untrack();
    await context.sync();
});
```

---

## Source

- /en-us/javascript/api/word/word.oleformat
