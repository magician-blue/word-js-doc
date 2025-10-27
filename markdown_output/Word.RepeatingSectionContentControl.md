# Word.RepeatingSectionContentControl

**Package:** `word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents the RepeatingSectionContentControl object.

## Properties

### allowInsertDeleteSection

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether users can add or remove sections from this repeating section content control by using the user interface.

#### Examples

**Example**: Prevent users from adding or removing sections in a repeating section content control through the UI

```typescript
await Word.run(async (context) => {
    // Get the first repeating section content control in the document
    const repeatingSections = context.document.contentControls.getByTypes([Word.ContentControlType.repeatingSectionItem]).getFirst();
    const repeatingSection = repeatingSections.parentContentControlOrNullObject as Word.RepeatingSectionContentControl;
    
    repeatingSection.load("allowInsertDeleteSection");
    await context.sync();
    
    // Disable the ability to insert or delete sections
    repeatingSection.allowInsertDeleteSection = false;
    
    await context.sync();
    console.log("Users can no longer add or remove sections from this repeating section");
});
```

---

### appearance

**Type:** `Word.ContentControlAppearance | "BoundingBox" | "Tags" | "Hidden"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the appearance of the content control.

#### Examples

**Example**: Set the appearance of a repeating section content control to show bounding box borders

```typescript
await Word.run(async (context) => {
    // Get the first repeating section content control in the document
    const repeatingSectionContentControls = context.document.contentControls.getByTypes([Word.ContentControlType.repeatingSectionItem]);
    const repeatingSectionContentControl = repeatingSectionContentControls.getFirstOrNullObject();
    
    // Load the content control
    repeatingSectionContentControl.load("appearance");
    await context.sync();
    
    // Set the appearance to show bounding box
    if (!repeatingSectionContentControl.isNullObject) {
        repeatingSectionContentControl.appearance = Word.ContentControlAppearance.boundingBox;
        await context.sync();
    }
});
```

---

### color

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the red-green-blue (RGB) value of the color of the content control. You can provide the value in the '#RRGGBB' format.

#### Examples

**Example**: Set the color of a repeating section content control to blue (#0000FF)

```typescript
await Word.run(async (context) => {
    // Get the first repeating section content control in the document
    const repeatingSectionContentControl = context.document.contentControls
        .getByTypes([Word.ContentControlType.repeatingSectionItem])
        .getFirst();
    
    // Set the color to blue
    repeatingSectionContentControl.color = "#0000FF";
    
    await context.sync();
});
```

---

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a repeating section content control to verify the connection is established before performing operations on the control.

```typescript
await Word.run(async (context) => {
    // Get the first repeating section content control in the document
    const repeatingSections = context.document.contentControls.getByTypes([Word.ContentControlType.repeatingSectionItem]);
    const repeatingSection = repeatingSections.getFirstOrNullObject();
    
    // Load the repeating section
    repeatingSection.load("id");
    await context.sync();
    
    if (!repeatingSection.isNullObject) {
        // Access the context property to verify connection
        const requestContext = repeatingSection.context;
        
        // Use the context to perform operations
        console.log("Request context is available:", requestContext !== null);
        console.log("Repeating section ID:", repeatingSection.id);
    }
});
```

---

### id

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns the identification for the content control.

#### Examples

**Example**: Get the ID of a repeating section content control and display it in the console for tracking purposes.

```typescript
await Word.run(async (context) => {
    // Get the first repeating section content control in the document
    const repeatingSectionContentControls = context.document.contentControls.getByTypes([Word.ContentControlType.repeatingSectionItem]);
    const repeatingSectionCC = repeatingSectionContentControls.getFirstOrNullObject();
    
    // Load the id property
    repeatingSectionCC.load("id");
    
    await context.sync();
    
    if (!repeatingSectionCC.isNullObject) {
        // Access and display the content control ID
        console.log("Repeating Section Content Control ID: " + repeatingSectionCC.id);
    } else {
        console.log("No repeating section content control found.");
    }
});
```

---

### isTemporary

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether to remove the content control from the active document when the user edits the contents of the control.

#### Examples

**Example**: Set a repeating section content control to be automatically removed when the user edits its contents

```typescript
await Word.run(async (context) => {
    // Get the first repeating section content control in the document
    const repeatingSectionContentControls = context.document.contentControls.getByTypes([Word.ContentControlType.repeatingSectionItem]);
    const repeatingSectionContentControl = repeatingSectionContentControls.getFirstOrNullObject() as Word.RepeatingSectionContentControl;
    
    repeatingSectionContentControl.load("isTemporary");
    await context.sync();
    
    // Set the content control to be temporary (removed when user edits it)
    repeatingSectionContentControl.isTemporary = true;
    
    await context.sync();
    console.log("Repeating section content control set to temporary");
});
```

---

### level

**Type:** `Word.ContentControlLevel | "Inline" | "Paragraph" | "Row" | "Cell"`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns the level of the content controlâwhether the content control surrounds text, paragraphs, table cells, or table rows; or if it is inline.

#### Examples

**Example**: Check the level of a repeating section content control and display it in the console to determine if it's inline, paragraph-level, row-level, or cell-level.

```typescript
await Word.run(async (context) => {
    // Get the first repeating section content control in the document
    const repeatingSections = context.document.contentControls.getByTypes([Word.ContentControlType.repeatingSectionItem]);
    const firstRepeatingSection = repeatingSections.getFirstOrNullObject() as Word.RepeatingSectionContentControl;
    
    // Load the level property
    firstRepeatingSection.load("level");
    
    await context.sync();
    
    if (!firstRepeatingSection.isNullObject) {
        // Display the level of the repeating section content control
        console.log(`Repeating section level: ${firstRepeatingSection.level}`);
    } else {
        console.log("No repeating section content control found.");
    }
});
```

---

### lockContentControl

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies if the content control is locked (can't be deleted). `true` means that the user can't delete it from the active document, `false` means it can be deleted.

#### Examples

**Example**: Lock a repeating section content control to prevent users from deleting it from the document

```typescript
await Word.run(async (context) => {
    // Get the first repeating section content control in the document
    const repeatingSections = context.document.contentControls.getByTypes([Word.ContentControlType.repeatingSectionItem]);
    const repeatingSection = repeatingSections.getFirstOrNullObject();
    
    repeatingSection.load("lockContentControl");
    await context.sync();
    
    if (!repeatingSection.isNullObject) {
        // Lock the content control to prevent deletion
        repeatingSection.lockContentControl = true;
        await context.sync();
        
        console.log("Repeating section content control is now locked");
    }
});
```

---

### lockContents

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies if the contents of the content control are locked (not editable). `true` means the user can't edit the contents, `false` means the contents are editable.

#### Examples

**Example**: Lock the contents of a repeating section content control to prevent users from editing the repeated items while still allowing them to add or remove sections.

```typescript
await Word.run(async (context) => {
    // Get the first repeating section content control in the document
    const repeatingSections = context.document.contentControls.getByTypes([Word.ContentControlType.repeatingSectionItem]);
    const repeatingSection = repeatingSections.getFirstOrNullObject() as Word.RepeatingSectionContentControl;
    
    repeatingSection.load("lockContents");
    await context.sync();
    
    // Lock the contents to prevent editing
    repeatingSection.lockContents = true;
    
    await context.sync();
    console.log("Repeating section contents are now locked");
});
```

---

### placeholderText

**Type:** `Word.BuildingBlock`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns a `BuildingBlock` object that represents the placeholder text for the content control.

#### Examples

**Example**: Get and display the placeholder text content from a repeating section content control by accessing its BuildingBlock object.

```typescript
await Word.run(async (context) => {
    // Get the first repeating section content control in the document
    const repeatingSections = context.document.contentControls.getByTypes([Word.ContentControlType.repeatingSectionItem]);
    const repeatingSection = repeatingSections.getFirstOrNullObject() as Word.RepeatingSectionContentControl;
    
    // Load the placeholder text building block
    const placeholderBlock = repeatingSection.placeholderText;
    placeholderBlock.load("value");
    
    await context.sync();
    
    // Display the placeholder text
    if (!repeatingSection.isNullObject) {
        console.log("Placeholder text: " + placeholderBlock.value);
    } else {
        console.log("No repeating section content control found.");
    }
});
```

---

### range

**Type:** `Word.Range`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets a `Range` object that represents the contents of the content control in the active document.

#### Examples

**Example**: Highlight all text within a repeating section content control by applying a yellow background color to its range.

```typescript
await Word.run(async (context) => {
    // Get the first repeating section content control in the document
    const repeatingSections = context.document.contentControls.getByTypes([Word.ContentControlType.repeatingSectionItem]);
    const repeatingSection = repeatingSections.getFirstOrNullObject() as Word.RepeatingSectionContentControl;
    
    // Get the range of the repeating section content control
    const range = repeatingSection.range;
    
    // Apply yellow highlight to the range
    range.font.highlightColor = "yellow";
    
    await context.sync();
});
```

---

### repeatingSectionItems

**Type:** `Word.RepeatingSectionItemCollection`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns the collection of repeating section items in this repeating section content control.

#### Examples

**Example**: Get the count of items in a repeating section content control and log each item's index to the console.

```typescript
await Word.run(async (context) => {
    // Get the first repeating section content control in the document
    const repeatingSections = context.document.contentControls.getByTypes([Word.ContentControlType.repeatingSection]);
    const repeatingSection = repeatingSections.getFirstOrNullObject() as Word.RepeatingSectionContentControl;
    
    // Get the collection of repeating section items
    const items = repeatingSection.repeatingSectionItems;
    items.load("items");
    
    await context.sync();
    
    // Log the count and index of each item
    console.log(`Total items in repeating section: ${items.items.length}`);
    items.items.forEach((item, index) => {
        console.log(`Item ${index + 1}`);
    });
});
```

---

### repeatingSectionItemTitle

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the name of the repeating section items used in the context menu associated with this repeating section content control.

#### Examples

**Example**: Set the repeating section item title to "Employee Record" so it appears in the context menu when users interact with the repeating section content control.

```typescript
await Word.run(async (context) => {
    // Get the first repeating section content control in the document
    const repeatingSectionCC = context.document.contentControls
        .getByTypes([Word.ContentControlType.repeatingSection])
        .getFirst() as Word.RepeatingSectionContentControl;
    
    // Set the item title that appears in the context menu
    repeatingSectionCC.repeatingSectionItemTitle = "Employee Record";
    
    await context.sync();
    
    console.log("Repeating section item title set to 'Employee Record'");
});
```

---

### showingPlaceholderText

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns whether the placeholder text for the content control is being displayed.

#### Examples

**Example**: Check if a repeating section content control is showing placeholder text and log the result to the console.

```typescript
await Word.run(async (context) => {
    // Get the first repeating section content control in the document
    const repeatingSectionContentControls = context.document.contentControls.getByTypes([Word.ContentControlType.repeatingSectionItem]);
    const firstRepeatingSectionCC = repeatingSectionContentControls.getFirstOrNullObject();
    
    firstRepeatingSectionCC.load("showingPlaceholderText");
    
    await context.sync();
    
    if (!firstRepeatingSectionCC.isNullObject) {
        if (firstRepeatingSectionCC.showingPlaceholderText) {
            console.log("The content control is showing placeholder text");
        } else {
            console.log("The content control has user-entered content");
        }
    }
});
```

---

### tag

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies a tag to identify the content control.

#### Examples

**Example**: Set a tag "employee-section" on a repeating section content control to identify it for later retrieval or processing.

```typescript
await Word.run(async (context) => {
    // Get the first repeating section content control in the document
    const repeatingSectionCC = context.document.contentControls
        .getByTypes([Word.ContentControlType.repeatingSectionItem])
        .getFirst();
    
    // Set the tag to identify this content control
    repeatingSectionCC.tag = "employee-section";
    
    await context.sync();
    console.log("Tag 'employee-section' has been set on the repeating section content control");
});
```

---

### title

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the title for the content control.

#### Examples

**Example**: Set the title of a repeating section content control to "Employee Records"

```typescript
await Word.run(async (context) => {
    const contentControls = context.document.contentControls.getByTag("repeatingSectionCC");
    contentControls.load("items");
    
    await context.sync();
    
    if (contentControls.items.length > 0) {
        const repeatingSectionCC = contentControls.items[0] as Word.RepeatingSectionContentControl;
        repeatingSectionCC.title = "Employee Records";
        
        await context.sync();
    }
});
```

---

### xmlapping

**Type:** `Word.XmlMapping`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns an `XmlMapping` object that represents the mapping of the content control to XML data in the data store of the document.

#### Examples

**Example**: Check if a repeating section content control has XML mapping configured and log the mapping details to the console.

```typescript
await Word.run(async (context) => {
    // Get the first repeating section content control in the document
    const repeatingSections = context.document.contentControls.getByTypes([Word.ContentControlType.repeatingSectionItem]);
    const repeatingSection = repeatingSections.getFirstOrNullObject() as Word.RepeatingSectionContentControl;
    
    // Load the XML mapping property
    repeatingSection.load("xmlMapping");
    
    await context.sync();
    
    if (!repeatingSection.isNullObject) {
        const xmlMapping = repeatingSection.xmlMapping;
        xmlMapping.load("customXmlPart, xpath");
        
        await context.sync();
        
        console.log("XML Mapping XPath:", xmlMapping.xpath);
        console.log("Has Custom XML Part:", xmlMapping.customXmlPart !== null);
    } else {
        console.log("No repeating section content control found.");
    }
});
```

---

## Methods

### copy

Copies the content control from the active document to the Clipboard.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Copy a repeating section content control to the clipboard so it can be pasted elsewhere in the document

```typescript
await Word.run(async (context) => {
    // Get the first repeating section content control in the document
    const repeatingSections = context.document.contentControls.getByTypes([Word.ContentControlType.repeatingSectionItem]);
    const firstRepeatingSection = repeatingSections.getFirstOrNullObject();
    
    firstRepeatingSection.load("id");
    await context.sync();
    
    if (!firstRepeatingSection.isNullObject) {
        // Copy the repeating section content control to clipboard
        firstRepeatingSection.copy();
        await context.sync();
        
        console.log("Repeating section content control copied to clipboard");
    } else {
        console.log("No repeating section content control found");
    }
});
```

---

### cut

Removes the content control from the active document and moves the content control to the Clipboard.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Remove a repeating section content control from the document and move it to the clipboard so it can be pasted elsewhere

```typescript
await Word.run(async (context) => {
    // Get the first repeating section content control in the document
    const repeatingSections = context.document.contentControls.getByTypes([Word.ContentControlType.repeatingSectionItem]);
    repeatingSections.load("items");
    
    await context.sync();
    
    if (repeatingSections.items.length > 0) {
        const firstRepeatingSection = repeatingSections.items[0] as Word.RepeatingSectionContentControl;
        
        // Cut the repeating section content control to clipboard
        firstRepeatingSection.cut();
        
        await context.sync();
        console.log("Repeating section content control cut to clipboard");
    }
});
```

---

### delete

**Kind:** `delete`

Deletes the content control and the contents of the content control.

#### Signature

**Parameters:**
- `deleteContents`: `boolean` (optional)
  Optional. Whether to delete the contents inside the control.

**Returns:** `void`

#### Examples

**Example**: Delete a repeating section content control and all its contents from the document

```typescript
await Word.run(async (context) => {
    // Get the first repeating section content control in the document
    const repeatingSections = context.document.contentControls.getByTypes([Word.ContentControlType.repeatingSectionItem]);
    const firstRepeatingSection = repeatingSections.getFirstOrNullObject();
    
    firstRepeatingSection.load("id");
    await context.sync();
    
    if (!firstRepeatingSection.isNullObject) {
        // Delete the repeating section and its contents
        firstRepeatingSection.delete(true);
        await context.sync();
        
        console.log("Repeating section content control deleted successfully");
    } else {
        console.log("No repeating section content control found");
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
  - `options`: `Word.Interfaces.RepeatingSectionContentControlLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.RepeatingSectionContentControl`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.RepeatingSectionContentControl`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.RepeatingSectionContentControl`

#### Examples

**Example**: Load and read the title property of the first repeating section content control in the document

```typescript
await Word.run(async (context) => {
    // Get the first repeating section content control
    const repeatingSections = context.document.contentControls.getByTypes([Word.ContentControlType.repeatingSectionItem]);
    const firstRepeatingSection = repeatingSections.getFirstOrNullObject() as Word.RepeatingSectionContentControl;
    
    // Load the title property
    firstRepeatingSection.load("title");
    
    // Sync to execute the load command
    await context.sync();
    
    // Check if the control exists and read the property
    if (!firstRepeatingSection.isNullObject) {
        console.log("Repeating section title: " + firstRepeatingSection.title);
    } else {
        console.log("No repeating section content control found");
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
  - `properties`: `Interfaces.RepeatingSectionContentControlUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.RepeatingSectionContentControl` (required)

  **Returns:** `void`

#### Examples

**Example**: Update multiple properties of a repeating section content control, setting both its title and appearance properties at once.

```typescript
await Word.run(async (context) => {
    // Get the first repeating section content control in the document
    const repeatingSectionCC = context.document.contentControls
        .getByTypes([Word.ContentControlType.repeatingSectionItem])
        .getFirst();
    
    // Set multiple properties at once using the set() method
    repeatingSectionCC.set({
        title: "Employee Records",
        appearance: Word.ContentControlAppearance.boundingBox,
        color: "blue"
    });
    
    await context.sync();
    console.log("Repeating section content control properties updated");
});
```

---

### setPlaceholderText

**Kind:** `configure`

Sets the placeholder text that displays in the content control until a user enters their own text.

#### Signature

**Parameters:**
- `options`: `Word.ContentControlPlaceholderOptions` (optional)
  Optional. The options for configuring the content control's placeholder text.

**Returns:** `void`

#### Examples

**Example**: Set placeholder text for a repeating section content control to guide users to add employee records

```typescript
await Word.run(async (context) => {
    // Get the first repeating section content control in the document
    const repeatingSections = context.document.contentControls.getByTypes([Word.ContentControlType.repeatingSectionItem]);
    const repeatingSection = repeatingSections.getFirstOrNullObject();
    
    repeatingSection.load("type");
    await context.sync();
    
    if (!repeatingSection.isNullObject) {
        // Set placeholder text for the repeating section
        repeatingSection.setPlaceholderText("Click here to add employee information");
        await context.sync();
    }
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.RepeatingSectionContentControl` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.RepeatingSectionContentControlData`) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.RepeatingSectionContentControlData`

#### Examples

**Example**: Serialize a repeating section content control to JSON format for logging or data export purposes.

```typescript
await Word.run(async (context) => {
    // Get the first repeating section content control in the document
    const repeatingSections = context.document.contentControls.getByTypes([Word.ContentControlType.repeatingSectionItem]);
    const repeatingSection = repeatingSections.getFirstOrNullObject();
    
    // Load properties to serialize
    repeatingSection.load("id,tag,title,type,text");
    
    await context.sync();
    
    if (!repeatingSection.isNullObject) {
        // Convert the repeating section content control to a plain JSON object
        const jsonData = repeatingSection.toJSON();
        
        // Log the JSON representation
        console.log("Repeating Section Content Control as JSON:");
        console.log(JSON.stringify(jsonData, null, 2));
        
        // You can now use this plain object for data export, storage, or comparison
        const exportData = {
            timestamp: new Date().toISOString(),
            controlData: jsonData
        };
        console.log("Export package:", exportData);
    }
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.RepeatingSectionContentControl`

#### Examples

**Example**: Track a repeating section content control across multiple sync calls to safely modify its properties without encountering "InvalidObjectPath" errors.

```typescript
await Word.run(async (context) => {
    // Get the first repeating section content control in the document
    const repeatingSections = context.document.contentControls.getByTypes([Word.ContentControlType.repeatingSectionItem]);
    const repeatingSection = repeatingSections.getFirstOrNullObject() as Word.RepeatingSectionContentControl;
    
    // Track the object for use across sync calls
    repeatingSection.track();
    
    await context.sync();
    
    // Now we can safely use the object after sync
    if (!repeatingSection.isNullObject) {
        repeatingSection.title = "Updated Repeating Section";
        await context.sync();
        
        // Continue working with the tracked object
        repeatingSection.tag = "tracked-section";
        await context.sync();
    }
    
    // Untrack when done to free up memory
    repeatingSection.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

#### Signature

**Returns:** `Word.RepeatingSectionContentControl`

#### Examples

**Example**: Process a repeating section content control and then release it from memory tracking to improve performance after you're done using it.

```typescript
await Word.run(async (context) => {
    // Get the first repeating section content control in the document
    const repeatingSectionCC = context.document.contentControls
        .getByTypes([Word.ContentControlType.repeatingSectionItem])
        .getFirstOrNullObject();
    
    // Load and track the object
    repeatingSectionCC.load("tag");
    await context.sync();
    
    if (!repeatingSectionCC.isNullObject) {
        // Work with the repeating section content control
        console.log("Repeating section tag: " + repeatingSectionCC.tag);
        
        // Untrack the object to free memory after we're done using it
        repeatingSectionCC.untrack();
        await context.sync();
    }
});
```

---

## Source

- https://learn.microsoft.com/en-us/javascript/api/word/word.repeatingsectioncontentcontrol
- https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets
