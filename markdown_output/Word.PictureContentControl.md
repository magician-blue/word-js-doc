# Word.PictureContentControl

**Package:** `word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents the PictureContentControl object.

## Properties

### appearance

**Type:** `Word.ContentControlAppearance | "BoundingBox" | "Tags" | "Hidden"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the appearance of the content control.

#### Examples

**Example**: Set a picture content control's appearance to show only a bounding box without tags

```typescript
await Word.run(async (context) => {
    // Get the first picture content control in the document
    const pictureContentControls = context.document.contentControls.getByTypes([Word.ContentControlType.picture]);
    const pictureContentControl = pictureContentControls.getFirst() as Word.PictureContentControl;
    
    // Set the appearance to BoundingBox
    pictureContentControl.appearance = Word.ContentControlAppearance.boundingBox;
    
    await context.sync();
});
```

---

### color

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the red-green-blue (RGB) value of the color of the content control. You can provide the value in the '#RRGGBB' format.

#### Examples

**Example**: Set the color of a picture content control to blue (#0000FF)

```typescript
await Word.run(async (context) => {
    // Get the first picture content control in the document
    const pictureContentControl = context.document.contentControls.getByTypes([Word.ContentControlType.picture]).getFirst();
    
    // Set the color to blue
    pictureContentControl.color = "#0000FF";
    
    await context.sync();
});
```

---

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a picture content control to verify the connection to the Office host application and log context information.

```typescript
await Word.run(async (context) => {
    // Get the first picture content control in the document
    const pictureContentControls = context.document.contentControls.getByTypes([Word.ContentControlType.picture]);
    const pictureContentControl = pictureContentControls.getFirst();
    
    // Load the picture content control
    pictureContentControl.load("id");
    await context.sync();
    
    // Access the request context from the picture content control
    const requestContext = pictureContentControl.context;
    
    // Verify the context is available and log information
    console.log("Request context is connected:", requestContext !== null);
    console.log("Picture content control ID:", pictureContentControl.id);
    
    // The context property allows you to perform operations on the same context
    // For example, you can use it to access the document through the control's context
    const document = requestContext.document;
    document.body.load("text");
    await requestContext.sync();
    
    console.log("Document accessed via picture control's context");
});
```

---

### id

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns the identification for the content control.

#### Examples

**Example**: Retrieve and display the unique identifier of a picture content control in the document.

```typescript
await Word.run(async (context) => {
    // Get the first picture content control in the document
    const pictureContentControls = context.document.contentControls.getByTypes([Word.ContentControlType.picture]);
    const pictureContentControl = pictureContentControls.getFirstOrNullObject();
    
    // Load the id property
    pictureContentControl.load("id");
    
    await context.sync();
    
    if (!pictureContentControl.isNullObject) {
        // Display the content control ID
        console.log("Picture Content Control ID: " + pictureContentControl.id);
    } else {
        console.log("No picture content control found in the document.");
    }
});
```

---

### isTemporary

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether to remove the content control from the active document when the user edits the contents of the control.

#### Examples

**Example**: Set a picture content control to be automatically removed from the document when the user edits its contents

```typescript
await Word.run(async (context) => {
    // Get the first picture content control in the document
    const pictureContentControls = context.document.contentControls.getByTypes([Word.ContentControlType.picture]);
    const pictureControl = pictureContentControls.getFirstOrNullObject();
    
    pictureControl.load("isTemporary");
    await context.sync();
    
    if (!pictureControl.isNullObject) {
        // Set the control to be temporary (removed when user edits it)
        pictureControl.isTemporary = true;
        await context.sync();
        
        console.log("Picture content control set to temporary");
    }
});
```

---

### level

**Type:** `Word.ContentControlLevel | "Inline" | "Paragraph" | "Row" | "Cell"`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns the level of the content control—whether the content control surrounds text, paragraphs, table cells, or table rows; or if it is inline.

#### Examples

**Example**: Check if a picture content control is inline or at paragraph level and display the level information in the console.

```typescript
await Word.run(async (context) => {
    // Get the first picture content control in the document
    const pictureContentControls = context.document.contentControls.getByTypes([Word.ContentControlType.picture]);
    const pictureContentControl = pictureContentControls.getFirstOrNullObject() as Word.PictureContentControl;
    
    // Load the level property
    pictureContentControl.load("level");
    
    await context.sync();
    
    if (!pictureContentControl.isNullObject) {
        // Display the content control level
        console.log(`Picture content control level: ${pictureContentControl.level}`);
        
        // You can also check the level programmatically
        if (pictureContentControl.level === Word.ContentControlLevel.inline) {
            console.log("This is an inline picture content control");
        } else if (pictureContentControl.level === Word.ContentControlLevel.paragraph) {
            console.log("This is a paragraph-level picture content control");
        }
    } else {
        console.log("No picture content control found in the document");
    }
});
```

---

### lockContentControl

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies if the content control is locked (can't be deleted). true means that the user can't delete it from the active document, false means it can be deleted.

#### Examples

**Example**: Lock a picture content control to prevent users from deleting it from the document

```typescript
await Word.run(async (context) => {
    // Get the first picture content control in the document
    const pictureContentControl = context.document.contentControls.getByTag("myPicture").getFirst();
    
    // Lock the content control to prevent deletion
    pictureContentControl.lockContentControl = true;
    
    await context.sync();
    console.log("Picture content control is now locked and cannot be deleted.");
});
```

---

### lockContents

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies if the contents of the content control are locked (not editable). true means the user can't edit the contents, false means the contents are editable.

#### Examples

**Example**: Lock the contents of a picture content control to prevent users from editing or replacing the image

```typescript
await Word.run(async (context) => {
    // Get the first picture content control in the document
    const pictureContentControls = context.document.contentControls.getByTypes([Word.ContentControlType.picture]);
    const pictureControl = pictureContentControls.getFirst();
    
    // Lock the contents to prevent editing
    pictureControl.lockContents = true;
    
    await context.sync();
    console.log("Picture content control contents are now locked.");
});
```

---

### placeholderText

**Type:** `Word.BuildingBlock`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns a BuildingBlock object that represents the placeholder text for the content control.

#### Examples

**Example**: Get and display the placeholder text content from a picture content control by accessing its BuildingBlock object properties.

```typescript
await Word.run(async (context) => {
    // Get the first picture content control in the document
    const pictureContentControls = context.document.contentControls.getByTypes([Word.ContentControlType.picture]);
    const pictureCC = pictureContentControls.getFirstOrNullObject();
    
    // Load the placeholder text BuildingBlock
    const placeholder = pictureCC.placeholderText;
    placeholder.load("value");
    
    await context.sync();
    
    if (!pictureCC.isNullObject) {
        console.log("Placeholder text: " + placeholder.value);
    } else {
        console.log("No picture content control found.");
    }
});
```

---

### range

**Type:** `Word.Range`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns a Range object that represents the contents of the content control in the active document.

#### Examples

**Example**: Get the text content from a picture content control by accessing its range and reading the text property.

```typescript
await Word.run(async (context) => {
    // Get the first picture content control in the document
    const pictureContentControl = context.document.contentControls.getByTag("myPicture").getFirst();
    
    // Get the range of the picture content control
    const range = pictureContentControl.range;
    
    // Load the text property of the range
    range.load("text");
    
    await context.sync();
    
    // Log the text content (if any) within the picture content control
    console.log("Content control range text: " + range.text);
});
```

---

### showingPlaceholderText

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns whether the placeholder text for the content control is being displayed.

#### Examples

**Example**: Check if a picture content control is showing placeholder text and log the result to the console.

```typescript
await Word.run(async (context) => {
    const pictureContentControls = context.document.contentControls.getByTypes([Word.ContentControlType.picture]);
    pictureContentControls.load("showingPlaceholderText");
    
    await context.sync();
    
    if (pictureContentControls.items.length > 0) {
        const pictureControl = pictureContentControls.items[0];
        console.log("Is showing placeholder text: " + pictureControl.showingPlaceholderText);
    }
});
```

---

### tag

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies a tag to identify the content control.

#### Examples

**Example**: Set a tag "employee-photo" on a picture content control to identify it for later retrieval or processing.

```typescript
await Word.run(async (context) => {
    // Get the first picture content control in the document
    const pictureContentControl = context.document.contentControls.getByTypes([Word.ContentControlType.picture]).getFirst();
    
    // Set a tag to identify this picture content control
    pictureContentControl.tag = "employee-photo";
    
    await context.sync();
    
    console.log("Tag 'employee-photo' has been set on the picture content control");
});
```

---

### title

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the title for the content control.

#### Examples

**Example**: Set the title of a picture content control to "Company Logo" to help identify it in the document.

```typescript
await Word.run(async (context) => {
    // Get the first picture content control in the document
    const pictureContentControl = context.document.contentControls.getByTypes([Word.ContentControlType.picture]).getFirst();
    
    // Set the title property
    pictureContentControl.title = "Company Logo";
    
    await context.sync();
});
```

---

### xmlMapping

**Type:** `Word.XmlMapping`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns an XmlMapping object that represents the mapping of the content control to XML data in the data store of the document.

#### Examples

**Example**: Check if a picture content control has XML mapping configured and display the mapping's XPath expression if it exists.

```typescript
await Word.run(async (context) => {
    // Get the first picture content control in the document
    const pictureContentControls = context.document.contentControls.getByTypes([Word.ContentControlType.picture]);
    const pictureCC = pictureContentControls.getFirstOrNullObject();
    
    // Load the xmlMapping property
    pictureCC.load("xmlMapping");
    
    await context.sync();
    
    if (!pictureCC.isNullObject) {
        const xmlMapping = pictureCC.xmlMapping;
        xmlMapping.load("xpath");
        
        await context.sync();
        
        if (xmlMapping.xpath) {
            console.log("XML Mapping XPath: " + xmlMapping.xpath);
        } else {
            console.log("No XML mapping configured for this picture content control");
        }
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

**Example**: Copy a picture content control to the clipboard so it can be pasted elsewhere in the document

```typescript
await Word.run(async (context) => {
    // Get the first picture content control in the document
    const pictureContentControls = context.document.contentControls.getByTypes([Word.ContentControlType.picture]);
    pictureContentControls.load("items");
    
    await context.sync();
    
    if (pictureContentControls.items.length > 0) {
        const pictureControl = pictureContentControls.items[0] as Word.PictureContentControl;
        
        // Copy the picture content control to the clipboard
        pictureControl.copy();
        
        await context.sync();
        console.log("Picture content control copied to clipboard");
    } else {
        console.log("No picture content controls found in the document");
    }
});
```

---

### cut

Removes the content control from the active document and moves the content control to the Clipboard.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Cut the first picture content control from the document and move it to the clipboard

```typescript
await Word.run(async (context) => {
    // Get the first picture content control in the document
    const pictureContentControls = context.document.contentControls.getByTypes([Word.ContentControlType.picture]);
    pictureContentControls.load("items");
    
    await context.sync();
    
    if (pictureContentControls.items.length > 0) {
        const firstPictureControl = pictureContentControls.items[0] as Word.PictureContentControl;
        
        // Cut the picture content control to clipboard
        firstPictureControl.cut();
        
        await context.sync();
        console.log("Picture content control cut to clipboard");
    } else {
        console.log("No picture content controls found");
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
  Optional. Decides whether to delete the contents of the content control.

**Returns:** `void`

#### Examples

**Example**: Delete the first picture content control in the document while keeping its image content in the document

```typescript
await Word.run(async (context) => {
    // Get the first picture content control in the document
    const pictureContentControls = context.document.contentControls.getByTypes([Word.ContentControlType.picture]);
    const firstPictureControl = pictureContentControls.getFirstOrNullObject();
    
    firstPictureControl.load("id");
    await context.sync();
    
    if (!firstPictureControl.isNullObject) {
        // Delete the content control but keep its contents (the image)
        firstPictureControl.delete(false);
        await context.sync();
        
        console.log("Picture content control deleted, image retained");
    } else {
        console.log("No picture content control found");
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
  - `options`: `Word.Interfaces.PictureContentControlLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.PictureContentControl`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.PictureContentControl`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.PictureContentControl`

#### Examples

**Example**: Load and display the image width and height properties of the first picture content control in the document.

```typescript
await Word.run(async (context) => {
    // Get the first picture content control in the document
    const pictureContentControls = context.document.contentControls.getByTypes([Word.ContentControlType.picture]);
    const pictureContentControl = pictureContentControls.getFirstOrNullObject();
    
    // Load specific properties of the picture content control
    pictureContentControl.load("image/width, image/height, title");
    
    // Sync to execute the load command
    await context.sync();
    
    // Check if picture content control exists and display properties
    if (!pictureContentControl.isNullObject) {
        console.log(`Picture Title: ${pictureContentControl.title}`);
        console.log(`Image Width: ${pictureContentControl.image.width}`);
        console.log(`Image Height: ${pictureContentControl.image.height}`);
    } else {
        console.log("No picture content control found in the document.");
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
  - `properties`: `Interfaces.PictureContentControlUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.PictureContentControl` (required)

  **Returns:** `void`

#### Examples

**Example**: Update multiple properties of a picture content control, including its title and tag, to organize and identify it within the document.

```typescript
await Word.run(async (context) => {
    // Get the first picture content control in the document
    const pictureContentControl = context.document.contentControls.getByTag("myPicture").getFirst();
    
    // Set multiple properties at once
    pictureContentControl.set({
        title: "Company Logo",
        tag: "companyLogo",
        cannotDelete: true,
        cannotEdit: false
    });
    
    await context.sync();
    console.log("Picture content control properties updated successfully");
});
```

---

### setPlaceholderText

Sets the placeholder text that displays in the content control until a user enters their own text.

#### Signature

**Parameters:**
- `options`: `Word.ContentControlPlaceholderOptions` (optional)
  Optional. The options for configuring the content control's placeholder text.

**Returns:** `void`

#### Examples

**Example**: Set placeholder text for a picture content control to guide users to insert a company logo

```typescript
await Word.run(async (context) => {
    // Get the first picture content control in the document
    const pictureContentControls = context.document.contentControls.getByTypes([Word.ContentControlType.picture]);
    const pictureControl = pictureContentControls.getFirst() as Word.PictureContentControl;
    
    // Set placeholder text to guide the user
    pictureControl.setPlaceholderText("Click here to insert your company logo");
    
    await context.sync();
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.PictureContentControl object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.PictureContentControlData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.PictureContentControlData`

#### Examples

**Example**: Get a JSON representation of a picture content control to log or store its properties for debugging purposes.

```typescript
await Word.run(async (context) => {
    // Get the first picture content control in the document
    const pictureContentControls = context.document.contentControls.getByTypes([Word.ContentControlType.picture]);
    const pictureControl = pictureContentControls.getFirstOrNullObject();
    
    // Load properties we want to include in the JSON output
    pictureControl.load("id,tag,title,appearance");
    
    await context.sync();
    
    if (!pictureControl.isNullObject) {
        // Convert the picture content control to a plain JavaScript object
        const jsonData = pictureControl.toJSON();
        
        // Now you can use the plain object (e.g., log it, store it, etc.)
        console.log("Picture Content Control Data:", JSON.stringify(jsonData, null, 2));
    } else {
        console.log("No picture content control found in the document.");
    }
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.PictureContentControl`

#### Examples

**Example**: Track a picture content control across multiple sync calls to update its properties without getting an InvalidObjectPath error

```typescript
await Word.run(async (context) => {
    // Get the first picture content control in the document
    const pictureContentControls = context.document.contentControls.getByTypes([Word.ContentControlType.picture]);
    context.load(pictureContentControls, "items");
    await context.sync();
    
    const pictureControl = pictureContentControls.items[0] as Word.PictureContentControl;
    
    // Track the object to use it across multiple sync calls
    pictureControl.track();
    
    // First sync - load initial properties
    pictureControl.load("title,tag");
    await context.sync();
    
    console.log("Current title:", pictureControl.title);
    
    // Second sync - modify properties (tracking prevents InvalidObjectPath error)
    pictureControl.title = "Updated Picture Title";
    pictureControl.tag = "tracked-picture";
    await context.sync();
    
    // Untrack when done to free up memory
    pictureControl.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.PictureContentControl`

#### Examples

**Example**: Release memory for a tracked picture content control after modifying its properties to prevent memory leaks in the host application.

```typescript
await Word.run(async (context) => {
    // Get the first picture content control in the document
    const pictureContentControl = context.document.contentControls.getByTag("myPicture").getFirstOrNullObject();
    
    // Track the object for change monitoring
    pictureContentControl.track();
    
    // Load and modify properties
    pictureContentControl.load("title");
    await context.sync();
    
    pictureContentControl.title = "Updated Picture";
    await context.sync();
    
    // Untrack the object to release memory after we're done using it
    pictureContentControl.untrack();
    await context.sync();
    
    console.log("Picture content control updated and memory released");
});
```

---

## Source

- /en-us/javascript/api/word
