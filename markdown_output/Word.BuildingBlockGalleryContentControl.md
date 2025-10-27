# BuildingBlockGalleryContentControl

**Package:** `word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents the BuildingBlockGalleryContentControl object.

## Properties

### appearance

**Type:** `Word.ContentControlAppearance | "BoundingBox" | "Tags" | "Hidden"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the appearance of the content control.

#### Examples

**Example**: Set the appearance of a building block gallery content control to show a bounding box around it

```typescript
await Word.run(async (context) => {
    // Get the first building block gallery content control in the document
    const contentControl = context.document.contentControls.getByType(Word.ContentControlType.buildingBlockGallery).getFirstOrNullObject();
    
    // Set the appearance to show a bounding box
    contentControl.appearance = Word.ContentControlAppearance.boundingBox;
    
    await context.sync();
});
```

---

### buildingBlockCategory

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the category for the building block content control.

#### Examples

**Example**: Set the building block category to "General" for a building block gallery content control

```typescript
await Word.run(async (context) => {
    // Get the first building block gallery content control in the document
    const contentControl = context.document.contentControls.getByTypes([Word.ContentControlType.buildingBlockGallery]).getFirst();
    
    // Set the building block category to "General"
    contentControl.buildingBlockCategory = "General";
    
    await context.sync();
    
    console.log("Building block category set to General");
});
```

---

### buildingBlockType

**Type:** `Word.BuildingBlockType | "QuickParts" | "CoverPage" | "Equations" | "Footers" | "Headers" | "PageNumber" | "Tables" | "Watermarks" | "AutoText" | "TextBox" | "PageNumberTop" | "PageNumberBottom" | "PageNumberPage" | "TableOfContents" | "CustomQuickParts" | "CustomCoverPage" | "CustomEquations" | "CustomFooters" | "CustomHeaders" | "CustomPageNumber" | "CustomTables" | "CustomWatermarks" | "CustomAutoText" | "CustomTextBox" | "CustomPageNumberTop" | "CustomPageNumberBottom" | "CustomPageNumberPage" | "CustomTableOfContents" | "Custom1" | "Custom2" | "Custom3" | "Custom4" | "Custom5" | "Bibliography" | "CustomBibliography"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies a BuildingBlockType value that represents the type of building block for the building block content control.

#### Examples

**Example**: Set a building block gallery content control to display only "Headers" type building blocks

```typescript
await Word.run(async (context) => {
    // Get the first building block gallery content control in the document
    const contentControls = context.document.contentControls.getByTypes([Word.ContentControlType.buildingBlockGallery]);
    const galleryControl = contentControls.getFirstOrNullObject() as Word.BuildingBlockGalleryContentControl;
    
    // Load the control to check if it exists
    galleryControl.load("id");
    await context.sync();
    
    if (!galleryControl.isNullObject) {
        // Set the building block type to Headers
        galleryControl.buildingBlockType = "Headers";
        
        await context.sync();
        console.log("Building block type set to Headers");
    }
});
```

---

### color

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the red-green-blue (RGB) value of the color of the content control. You can provide the value in the '#RRGGBB' format.

#### Examples

**Example**: Set the color of a building block gallery content control to blue (#0000FF)

```typescript
await Word.run(async (context) => {
    // Get the first building block gallery content control in the document
    const contentControl = context.document.contentControls.getByType(Word.ContentControlType.buildingBlockGallery).getFirstOrNullObject();
    
    // Set the color to blue
    contentControl.color = "#0000FF";
    
    await context.sync();
});
```

---

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a building block gallery content control to verify the connection to the Word host application and log its API version.

```typescript
await Word.run(async (context) => {
    // Get the first building block gallery content control in the document
    const contentControls = context.document.contentControls.getByTypes([Word.ContentControlType.buildingBlockGallery]);
    const galleryControl = contentControls.getFirstOrNullObject() as Word.BuildingBlockGalleryContentControl;
    
    await context.sync();
    
    if (!galleryControl.isNullObject) {
        // Access the request context from the content control
        const requestContext = galleryControl.context;
        
        // Use the context to check the API version
        console.log("Connected to Word host application");
        console.log("API version:", requestContext.diagnostics?.host);
        
        // The context property connects the add-in process to Office host
        // It's the same context object used for all operations
        console.log("Context is valid:", requestContext !== null);
    }
});
```

---

### id

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the identification for the content control.

#### Examples

**Example**: Get and display the unique identifier of a building block gallery content control in the document.

```typescript
await Word.run(async (context) => {
    // Get the first building block gallery content control in the document
    const contentControls = context.document.contentControls.getByTypes([Word.ContentControlType.buildingBlockGallery]);
    const galleryControl = contentControls.getFirstOrNullObject() as Word.BuildingBlockGalleryContentControl;
    
    // Load the id property
    galleryControl.load("id");
    
    await context.sync();
    
    if (!galleryControl.isNullObject) {
        console.log("Building block gallery content control ID: " + galleryControl.id);
    } else {
        console.log("No building block gallery content control found.");
    }
});
```

---

### isTemporary

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether to remove the content control from the active document when the user edits the contents of the control.

#### Examples

**Example**: Set a building block gallery content control to be automatically removed when the user edits its contents

```typescript
await Word.run(async (context) => {
    // Get the first building block gallery content control in the document
    const contentControls = context.document.contentControls.getByTypes([Word.ContentControlType.buildingBlockGallery]);
    const buildingBlockGallery = contentControls.getFirstOrNullObject() as Word.BuildingBlockGalleryContentControl;
    
    buildingBlockGallery.load("isTemporary");
    await context.sync();
    
    // Set the control to be temporary (removed when user edits it)
    buildingBlockGallery.isTemporary = true;
    
    await context.sync();
    console.log("Building block gallery control will be removed when edited");
});
```

---

### level

**Type:** `Word.ContentControlLevel | "Inline" | "Paragraph" | "Row" | "Cell"`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the level of the content control—whether the content control surrounds text, paragraphs, table cells, or table rows; or if it is inline.

#### Examples

**Example**: Check if a building block gallery content control is inline or at paragraph level and log the result to the console.

```typescript
await Word.run(async (context) => {
    const buildingBlockGalleryContentControl = context.document.contentControls.getByTag("myGallery").getFirst();
    buildingBlockGalleryContentControl.load("level");
    
    await context.sync();
    
    console.log(`Content control level: ${buildingBlockGalleryContentControl.level}`);
    
    if (buildingBlockGalleryContentControl.level === Word.ContentControlLevel.inline) {
        console.log("This is an inline content control");
    } else if (buildingBlockGalleryContentControl.level === Word.ContentControlLevel.paragraph) {
        console.log("This is a paragraph-level content control");
    }
});
```

---

### lockContentControl

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies if the content control is locked (can't be deleted). true means that the user can't delete it from the active document, false means it can be deleted.

#### Examples

**Example**: Lock a building block gallery content control to prevent users from deleting it from the document

```typescript
await Word.run(async (context) => {
    // Get the first building block gallery content control in the document
    const contentControl = context.document.contentControls.getByType(Word.ContentControlType.buildingBlockGallery).getFirstOrNullObject();
    
    contentControl.load("lockContentControl");
    await context.sync();
    
    if (!contentControl.isNullObject) {
        // Lock the content control to prevent deletion
        contentControl.lockContentControl = true;
        
        await context.sync();
        console.log("Building block gallery content control is now locked");
    }
});
```

---

### lockContents

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies if the contents of the content control are locked (not editable). true means the user can't edit the contents, false means the contents are editable.

#### Examples

**Example**: Lock the contents of a building block gallery content control to prevent users from editing it

```typescript
await Word.run(async (context) => {
    // Get the first building block gallery content control in the document
    const contentControl = context.document.contentControls.getByType(Word.ContentControlType.buildingBlockGallery).getFirstOrNullObject();
    
    contentControl.load("lockContents");
    await context.sync();
    
    if (!contentControl.isNullObject) {
        // Lock the contents to prevent editing
        contentControl.lockContents = true;
        
        await context.sync();
        console.log("Building block gallery content control contents are now locked");
    }
});
```

---

### placeholderText

**Type:** `Word.BuildingBlock`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns a BuildingBlock object that represents the placeholder text for the content control.

#### Examples

**Example**: Get and display the placeholder text content from a building block gallery content control

```typescript
await Word.run(async (context) => {
    // Get the first building block gallery content control in the document
    const contentControls = context.document.contentControls.getByTypes([Word.ContentControlType.buildingBlockGallery]);
    const galleryControl = contentControls.getFirstOrNullObject() as Word.BuildingBlockGalleryContentControl;
    
    // Load the placeholder text building block
    galleryControl.load("placeholderText");
    const placeholder = galleryControl.placeholderText;
    placeholder.load("value");
    
    await context.sync();
    
    if (!galleryControl.isNullObject) {
        console.log("Placeholder text: " + placeholder.value);
    } else {
        console.log("No building block gallery content control found.");
    }
});
```

---

### range

**Type:** `Word.Range`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns a Range object that represents the contents of the content control in the active document.

#### Examples

**Example**: Get the text content from a building block gallery content control by accessing its range

```typescript
await Word.run(async (context) => {
    // Get the first building block gallery content control in the document
    const contentControls = context.document.contentControls.getByTypes([Word.ContentControlType.buildingBlockGallery]);
    const galleryControl = contentControls.getFirstOrNullObject() as Word.BuildingBlockGalleryContentControl;
    
    // Get the range of the content control
    const range = galleryControl.range;
    range.load("text");
    
    await context.sync();
    
    // Use the range to access the content
    console.log("Content control text: " + range.text);
});
```

---

### showingPlaceholderText

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets if the placeholder text for the content control is being displayed.

#### Examples

**Example**: Check if a building block gallery content control is currently showing its placeholder text and log the result to the console.

```typescript
await Word.run(async (context) => {
    const contentControls = context.document.contentControls.getByTypes([Word.ContentControlType.buildingBlockGallery]);
    contentControls.load("items");
    
    await context.sync();
    
    if (contentControls.items.length > 0) {
        const galleryControl = contentControls.items[0] as Word.BuildingBlockGalleryContentControl;
        galleryControl.load("showingPlaceholderText");
        
        await context.sync();
        
        console.log(`Placeholder text is showing: ${galleryControl.showingPlaceholderText}`);
    }
});
```

---

### tag

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies a tag to identify the content control.

#### Examples

**Example**: Set a tag "product-catalog" on a building block gallery content control to identify it for later retrieval or manipulation.

```typescript
await Word.run(async (context) => {
    const contentControl = context.document.contentControls.getFirst();
    contentControl.load("type");
    await context.sync();
    
    if (contentControl.type === Word.ContentControlType.buildingBlockGallery) {
        const galleryControl = contentControl as Word.BuildingBlockGalleryContentControl;
        galleryControl.tag = "product-catalog";
        await context.sync();
    }
});
```

---

### title

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the title for the content control.

#### Examples

**Example**: Set the title of a building block gallery content control to "Document Templates"

```typescript
await Word.run(async (context) => {
    const contentControl = context.document.contentControls.getByTag("myGallery").getFirst();
    contentControl.title = "Document Templates";
    
    await context.sync();
});
```

---

### xmlMapping

**Type:** `Word.XmlMapping`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns an XmlMapping object that represents the mapping of the content control to XML data in the data store of the document.

#### Examples

**Example**: Check if a building block gallery content control has XML mapping configured and log the mapping details to the console.

```typescript
await Word.run(async (context) => {
    // Get the first building block gallery content control in the document
    const contentControls = context.document.contentControls.getByTypes([Word.ContentControlType.buildingBlockGallery]);
    const buildingBlockGallery = contentControls.getFirstOrNullObject() as Word.BuildingBlockGalleryContentControl;
    
    // Get the XML mapping for this content control
    const xmlMapping = buildingBlockGallery.xmlMapping;
    xmlMapping.load("isMapped, xpath, customXmlPart");
    
    await context.sync();
    
    if (!buildingBlockGallery.isNullObject) {
        console.log("XML Mapping Status:", xmlMapping.isMapped);
        console.log("XPath:", xmlMapping.xpath);
        console.log("Has Custom XML Part:", xmlMapping.customXmlPart !== null);
    } else {
        console.log("No building block gallery content control found.");
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

**Example**: Copy a building block gallery content control to the clipboard so it can be pasted elsewhere in the document

```typescript
await Word.run(async (context) => {
    // Get the first building block gallery content control in the document
    const contentControls = context.document.contentControls.getByTypes([Word.ContentControlType.buildingBlockGallery]);
    contentControls.load("items");
    
    await context.sync();
    
    if (contentControls.items.length > 0) {
        const buildingBlockGallery = contentControls.items[0] as Word.BuildingBlockGalleryContentControl;
        
        // Copy the content control to the clipboard
        buildingBlockGallery.copy();
        
        await context.sync();
        console.log("Building block gallery content control copied to clipboard");
    } else {
        console.log("No building block gallery content control found");
    }
});
```

---

### cut

Removes the content control from the active document and moves the content control to the Clipboard.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Remove a building block gallery content control from the document and place it on the clipboard so it can be pasted elsewhere

```typescript
await Word.run(async (context) => {
    // Get the first building block gallery content control in the document
    const contentControls = context.document.contentControls.getByTypes([Word.ContentControlType.buildingBlockGallery]);
    contentControls.load("items");
    
    await context.sync();
    
    if (contentControls.items.length > 0) {
        const buildingBlockGallery = contentControls.items[0] as Word.BuildingBlockGalleryContentControl;
        
        // Cut the content control to the clipboard
        buildingBlockGallery.cut();
        
        await context.sync();
        console.log("Building block gallery content control cut to clipboard");
    } else {
        console.log("No building block gallery content control found");
    }
});
```

---

### delete

**Kind:** `delete`

Deletes the content control and optionally its contents.

#### Signature

**Parameters:**
- `deleteContents`: `boolean` (optional)
  Optional. Whether to delete the contents inside the control.

**Returns:** `void`

#### Examples

**Example**: Delete a building block gallery content control while preserving its contents in the document

```typescript
await Word.run(async (context) => {
    // Get the first building block gallery content control in the document
    const contentControls = context.document.contentControls.getByTypes([Word.ContentControlType.buildingBlockGallery]);
    contentControls.load("items");
    
    await context.sync();
    
    if (contentControls.items.length > 0) {
        const galleryControl = contentControls.items[0] as Word.BuildingBlockGalleryContentControl;
        
        // Delete the control but keep its contents (false = don't delete contents)
        galleryControl.delete(false);
        
        await context.sync();
        console.log("Building block gallery content control deleted, contents preserved");
    }
});
```

---

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.BuildingBlockGalleryContentControlLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.BuildingBlockGalleryContentControl`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.BuildingBlockGalleryContentControl`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    select is a comma-delimited string that specifies the properties to load, and expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.BuildingBlockGalleryContentControl`

#### Examples

**Example**: Load and display the title and type properties of the first building block gallery content control in the document.

```typescript
await Word.run(async (context) => {
    // Get the first building block gallery content control
    const contentControls = context.document.contentControls.getByTypes([Word.ContentControlType.buildingBlockGallery]);
    const galleryControl = contentControls.getFirstOrNullObject() as Word.BuildingBlockGalleryContentControl;
    
    // Load specific properties
    galleryControl.load("title, type");
    
    // Sync to execute the load command
    await context.sync();
    
    // Check if control exists and display properties
    if (!galleryControl.isNullObject) {
        console.log(`Title: ${galleryControl.title}`);
        console.log(`Type: ${galleryControl.type}`);
    } else {
        console.log("No building block gallery content control found.");
    }
});
```

---

### set

**Kind:** `write`

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `Interfaces.BuildingBlockGalleryContentControlUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.BuildingBlockGalleryContentControl` (required)

  **Returns:** `void`

#### Examples

**Example**: Configure a building block gallery content control by setting multiple properties at once, including its title, appearance, and placeholder text.

```typescript
await Word.run(async (context) => {
    // Get the first building block gallery content control in the document
    const contentControl = context.document.contentControls.getFirstOrNullObject();
    contentControl.load("type");
    await context.sync();
    
    if (contentControl.type === Word.ContentControlType.buildingBlockGallery) {
        const galleryControl = contentControl as Word.BuildingBlockGalleryContentControl;
        
        // Set multiple properties at once using the set() method
        galleryControl.set({
            title: "Document Parts Gallery",
            appearance: Word.ContentControlAppearance.boundingBox,
            color: "#0078D4",
            placeholderText: "Select a building block to insert"
        });
        
        await context.sync();
        console.log("Building block gallery content control properties updated");
    }
});
```

---

### setPlaceholderText

Sets the placeholder text that displays in the content control until a user enters their own text.

#### Signature

**Parameters:**
- `options`: `Word.ContentControlPlaceholderOptions` (optional)
  The options for configuring the content control's placeholder text.

**Returns:** `void`

#### Examples

**Example**: Set placeholder text for a building block gallery content control to guide users to select a cover page template

```typescript
await Word.run(async (context) => {
    // Get the first building block gallery content control in the document
    const contentControls = context.document.contentControls.getByTypes([Word.ContentControlType.buildingBlockGallery]);
    const galleryControl = contentControls.getFirstOrNullObject();
    
    galleryControl.load("id");
    await context.sync();
    
    if (!galleryControl.isNullObject) {
        // Set placeholder text to guide the user
        galleryControl.setPlaceholderText("Click here to select a cover page template");
        await context.sync();
        
        console.log("Placeholder text set successfully");
    }
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.BuildingBlockGalleryContentControl object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.BuildingBlockGalleryContentControlData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.BuildingBlockGalleryContentControlData`

#### Examples

**Example**: Serialize a building block gallery content control to JSON format to log or store its properties

```typescript
await Word.run(async (context) => {
    // Get the first building block gallery content control in the document
    const contentControls = context.document.contentControls.getByTypes([Word.ContentControlType.buildingBlockGallery]);
    const galleryControl = contentControls.getFirstOrNullObject() as Word.BuildingBlockGalleryContentControl;
    
    // Load properties to include in JSON output
    galleryControl.load("id,tag,title,buildingBlockType,buildingBlockCategory");
    
    await context.sync();
    
    if (!galleryControl.isNullObject) {
        // Convert the control to a plain JavaScript object
        const jsonData = galleryControl.toJSON();
        
        // Now you can use the plain object (e.g., log it, store it, etc.)
        console.log("Building Block Gallery Control Data:", JSON.stringify(jsonData, null, 2));
    }
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.BuildingBlockGalleryContentControl`

#### Examples

**Example**: Track a building block gallery content control to maintain its reference across multiple sync calls and prevent "InvalidObjectPath" errors when accessing it later in the batch operation.

```typescript
await Word.run(async (context) => {
    // Get the first building block gallery content control in the document
    const contentControls = context.document.contentControls.getByTypes([Word.ContentControlType.buildingBlockGallery]);
    context.load(contentControls, "items");
    await context.sync();

    if (contentControls.items.length > 0) {
        const galleryControl = contentControls.items[0] as Word.BuildingBlockGalleryContentControl;
        
        // Track the object to use it across multiple sync calls
        galleryControl.track();
        
        await context.sync();
        
        // Now we can safely use the object in subsequent operations
        galleryControl.title = "Updated Gallery";
        await context.sync();
        
        // Access properties again without InvalidObjectPath error
        context.load(galleryControl, "title");
        await context.sync();
        console.log("Gallery title: " + galleryControl.title);
        
        // Untrack when done to free up memory
        galleryControl.untrack();
    }
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.BuildingBlockGalleryContentControl`

#### Examples

**Example**: Insert a building block gallery content control, use it to perform operations, then untrack it to free memory when done.

```typescript
await Word.run(async (context) => {
    // Insert a building block gallery content control
    const contentControl = context.document.body.insertContentControl(
        Word.ContentControlType.buildingBlockGallery
    );
    contentControl.load("id");
    
    await context.sync();
    
    // Perform operations with the content control
    console.log("Content control created with ID: " + contentControl.id);
    
    // Untrack the object to release memory
    contentControl.untrack();
    
    await context.sync();
    
    console.log("Content control untracked and memory released");
});
```

---

## Source

- /en-us/javascript/api/word
- /en-us/javascript/api/requirement-sets/word/word-api-requirement-sets
