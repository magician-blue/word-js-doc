# GroupContentControl

**Package:** `word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents the GroupContentControl object.

## Properties

### appearance

**Type:** `Word.ContentControlAppearance | "BoundingBox" | "Tags" | "Hidden"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the appearance of the content control.

#### Examples

**Example**: Set a group content control's appearance to show bounding box borders instead of tags

```typescript
await Word.run(async (context) => {
    // Get the first group content control in the document
    const groupContentControl = context.document.contentControls.getFirstOrNullObject();
    groupContentControl.load("type");
    
    await context.sync();
    
    if (groupContentControl.isNullObject) {
        console.log("No content control found");
        return;
    }
    
    // Set the appearance to show a bounding box
    groupContentControl.appearance = Word.ContentControlAppearance.boundingBox;
    // Or use the string literal: groupContentControl.appearance = "BoundingBox";
    
    await context.sync();
    console.log("Group content control appearance set to bounding box");
});
```

---

### color

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the red-green-blue (RGB) value of the color of the content control. You can provide the value in the '#RRGGBB' format.

#### Examples

**Example**: Set the color of a group content control to blue (#0000FF)

```typescript
await Word.run(async (context) => {
    // Get the first group content control in the document
    const groupContentControl = context.document.contentControls.getFirstOrNullObject();
    groupContentControl.load("type");
    
    await context.sync();
    
    if (groupContentControl.isNullObject) {
        console.log("No content control found");
        return;
    }
    
    // Set the color to blue
    groupContentControl.color = "#0000FF";
    
    await context.sync();
});
```

---

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access a group content control and use its context property to load and read the control's title property.

```typescript
await Word.run(async (context) => {
    // Get the first group content control in the document
    const groupContentControl = context.document.contentControls.getFirstOrNullObject();
    
    // Access the context property to load properties
    groupContentControl.context.load(groupContentControl, "title");
    
    await groupContentControl.context.sync();
    
    if (!groupContentControl.isNullObject) {
        console.log("Group Content Control Title: " + groupContentControl.title);
    }
});
```

---

### id

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns the identification for the content control.

#### Examples

**Example**: Retrieve and display the unique identifier of a group content control in the document.

```typescript
await Word.run(async (context) => {
    // Get the first group content control in the document
    const groupContentControl = context.document.contentControls.getFirstOrNullObject();
    groupContentControl.load("id, type");
    
    await context.sync();
    
    if (!groupContentControl.isNullObject && groupContentControl.type === "Group") {
        // Get the ID of the group content control
        const controlId = groupContentControl.id;
        console.log("Group Content Control ID: " + controlId);
    }
});
```

---

### isTemporary

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether to remove the content control from the active document when the user edits the contents of the control.

#### Examples

**Example**: Mark a group content control as temporary so it will be automatically removed when the user edits its contents

```typescript
await Word.run(async (context) => {
    // Get the first group content control in the document
    const groupContentControl = context.document.contentControls.getFirst();
    groupContentControl.load("type");
    
    await context.sync();
    
    // Verify it's a group content control
    if (groupContentControl.type === Word.ContentControlType.group) {
        const groupCC = groupContentControl as Word.GroupContentControl;
        
        // Set the control to be temporary (will be removed when user edits it)
        groupCC.isTemporary = true;
        
        await context.sync();
        console.log("Group content control marked as temporary");
    }
});
```

---

### level

**Type:** `Word.ContentControlLevel | "Inline" | "Paragraph" | "Row" | "Cell"`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the level of the content control—whether the content control surrounds text, paragraphs, table cells, or table rows; or if it is inline.

#### Examples

**Example**: Check the level of a group content control and display different messages based on whether it's inline, paragraph-level, row-level, or cell-level.

```typescript
await Word.run(async (context) => {
    // Get the first group content control in the document
    const groupContentControl = context.document.contentControls.getFirstOrNullObject();
    groupContentControl.load("level, type");
    
    await context.sync();
    
    if (!groupContentControl.isNullObject && groupContentControl.type === "Group") {
        const level = groupContentControl.level;
        
        switch (level) {
            case Word.ContentControlLevel.inline:
                console.log("This group content control is inline with text.");
                break;
            case Word.ContentControlLevel.paragraph:
                console.log("This group content control surrounds entire paragraphs.");
                break;
            case Word.ContentControlLevel.row:
                console.log("This group content control surrounds table rows.");
                break;
            case Word.ContentControlLevel.cell:
                console.log("This group content control surrounds table cells.");
                break;
        }
    }
});
```

---

### lockContentControl

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies if the content control is locked (can't be deleted). true means that the user can't delete it from the active document, false means it can be deleted.

#### Examples

**Example**: Lock a group content control to prevent users from deleting it from the document

```typescript
await Word.run(async (context) => {
    // Get the first group content control in the document
    const groupContentControl = context.document.contentControls.getFirstOrNullObject();
    groupContentControl.load("type");
    
    await context.sync();
    
    if (groupContentControl.isNullObject) {
        console.log("No content control found");
        return;
    }
    
    // Lock the content control to prevent deletion
    groupContentControl.lockContentControl = true;
    
    await context.sync();
    console.log("Group content control is now locked and cannot be deleted");
});
```

---

### lockContents

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies if the contents of the content control are locked (not editable). true means the user can't edit the contents, false means the contents are editable.

#### Examples

**Example**: Lock the contents of a group content control to prevent users from editing the grouped items while still allowing the entire group to be deleted or moved.

```typescript
await Word.run(async (context) => {
    // Get the first group content control in the document
    const groupContentControl = context.document.contentControls.getFirstOrNullObject();
    groupContentControl.load("type");
    
    await context.sync();
    
    // Verify it's a group content control and lock its contents
    if (groupContentControl.type === Word.ContentControlType.group) {
        groupContentControl.lockContents = true;
    }
    
    await context.sync();
    
    console.log("Group content control contents are now locked");
});
```

---

### placeholderText

**Type:** `Word.BuildingBlock`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns a BuildingBlock object that represents the placeholder text for the content control.

#### Examples

**Example**: Get and display the placeholder text content from a group content control by accessing its BuildingBlock properties.

```typescript
await Word.run(async (context) => {
    // Get the first group content control in the document
    const groupContentControl = context.document.contentControls.getFirstOrNullObject();
    groupContentControl.load("type");
    
    await context.sync();
    
    if (groupContentControl.isNullObject) {
        console.log("No content control found");
        return;
    }
    
    // Access the placeholder text BuildingBlock
    const placeholderBlock = groupContentControl.placeholderText;
    placeholderBlock.load("value");
    
    await context.sync();
    
    // Display the placeholder text value
    console.log("Placeholder text: " + placeholderBlock.value);
});
```

---

### range

**Type:** `Word.Range`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets a Range object that represents the contents of the content control in the active document.

#### Examples

**Example**: Get the text content from a group content control and display it in the console

```typescript
await Word.run(async (context) => {
    // Get the first group content control in the document
    const groupContentControl = context.document.contentControls.getFirstOrNullObject();
    groupContentControl.load("type");
    
    await context.sync();
    
    if (groupContentControl.type === Word.ContentControlType.group) {
        // Get the range of the group content control
        const range = groupContentControl.range;
        range.load("text");
        
        await context.sync();
        
        // Display the text content
        console.log("Group content control text: " + range.text);
    }
});
```

---

### showingPlaceholderText

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns whether the placeholder text for the content control is being displayed.

#### Examples

**Example**: Check if a group content control is displaying placeholder text and log the result to the console.

```typescript
await Word.run(async (context) => {
    const groupContentControl = context.document.contentControls.getFirst();
    groupContentControl.load("showingPlaceholderText");
    
    await context.sync();
    
    console.log("Is showing placeholder text: " + groupContentControl.showingPlaceholderText);
});
```

---

### tag

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies a tag to identify the content control.

#### Examples

**Example**: Set a tag "employee-info" on a group content control to identify it for later retrieval and processing.

```typescript
await Word.run(async (context) => {
    // Get the first group content control in the document
    const groupContentControl = context.document.contentControls.getFirstOrNullObject();
    groupContentControl.load("type");
    
    await context.sync();
    
    if (!groupContentControl.isNullObject && groupContentControl.type === "Group") {
        // Set a tag to identify this group content control
        groupContentControl.tag = "employee-info";
        
        await context.sync();
        console.log("Tag 'employee-info' has been set on the group content control");
    }
});
```

---

### title

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the title for the content control.

#### Examples

**Example**: Set the title of a group content control to "Employee Information Section"

```typescript
await Word.run(async (context) => {
    // Get the first group content control in the document
    const groupContentControl = context.document.contentControls.getByTag("employeeGroup").getFirst();
    
    // Set the title property
    groupContentControl.title = "Employee Information Section";
    
    await context.sync();
});
```

---

### xmlMapping

**Type:** `Word.XmlMapping`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets an XmlMapping object that represents the mapping of the content control to XML data in the data store of the document.

#### Examples

**Example**: Check if a group content control has an XML mapping and display its namespace URI and XPath expression.

```typescript
await Word.run(async (context) => {
    // Get the first group content control in the document
    const groupContentControls = context.document.contentControls.getByTypes([Word.ContentControlType.group]);
    const groupContentControl = groupContentControls.getFirst() as Word.GroupContentControl;
    
    // Get the XML mapping for the group content control
    const xmlMapping = groupContentControl.xmlMapping;
    
    // Load the XML mapping properties
    xmlMapping.load(["isMapped", "namespaceUri", "xpath"]);
    
    await context.sync();
    
    // Check and display the XML mapping information
    if (xmlMapping.isMapped) {
        console.log("Namespace URI: " + xmlMapping.namespaceUri);
        console.log("XPath: " + xmlMapping.xpath);
    } else {
        console.log("This group content control is not mapped to XML data.");
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

**Example**: Copy a group content control to the clipboard so it can be pasted elsewhere in the document

```typescript
await Word.run(async (context) => {
    // Get the first group content control in the document
    const groupContentControls = context.document.contentControls.getByTypes([Word.ContentControlType.group]);
    groupContentControls.load("items");
    
    await context.sync();
    
    if (groupContentControls.items.length > 0) {
        const groupControl = groupContentControls.items[0] as Word.GroupContentControl;
        
        // Copy the group content control to the clipboard
        groupControl.copy();
        
        await context.sync();
        console.log("Group content control copied to clipboard");
    } else {
        console.log("No group content controls found in the document");
    }
});
```

---

### cut

Removes the content control from the active document and moves the content control to the Clipboard.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Cut a group content control from the document and move it to the clipboard so it can be pasted elsewhere

```typescript
await Word.run(async (context) => {
    // Get the first group content control in the document
    const groupContentControls = context.document.contentControls.getByTypes([Word.ContentControlType.group]);
    groupContentControls.load("items");
    
    await context.sync();
    
    if (groupContentControls.items.length > 0) {
        const groupControl = groupContentControls.items[0] as Word.GroupContentControl;
        
        // Cut the group content control to clipboard
        groupControl.cut();
        
        await context.sync();
        console.log("Group content control has been cut to clipboard");
    } else {
        console.log("No group content controls found in the document");
    }
});
```

---

### delete

**Kind:** `delete`

Deletes the content control and optionally its contents.

#### Signature

**Parameters:**
- `deleteContents`: `boolean` (required)
  Optional. Whether to delete the contents inside the control.

**Returns:** `void`

#### Examples

**Example**: Delete a group content control while preserving its contents in the document

```typescript
await Word.run(async (context) => {
    // Get the first group content control in the document
    const groupContentControl = context.document.contentControls.getFirstOrNullObject();
    groupContentControl.load("type");
    
    await context.sync();
    
    // Check if it exists and is a group type
    if (!groupContentControl.isNullObject && groupContentControl.type === Word.ContentControlType.group) {
        // Delete the group content control but keep its contents
        groupContentControl.delete(false);
        
        await context.sync();
        console.log("Group content control deleted, contents preserved");
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
  - `options`: `Word.Interfaces.GroupContentControlLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.GroupContentControl`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.GroupContentControl`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.GroupContentControl`

#### Examples

**Example**: Load and display the ID and appearance properties of the first group content control in the document

```typescript
await Word.run(async (context) => {
    // Get the first group content control in the document
    const groupContentControl = context.document.contentControls.getFirstOrNullObject();
    
    // Load specific properties of the group content control
    groupContentControl.load("id, appearance, cannotDelete");
    
    // Synchronize the document state
    await context.sync();
    
    // Check if the content control exists and display its properties
    if (!groupContentControl.isNullObject) {
        console.log(`Group Content Control ID: ${groupContentControl.id}`);
        console.log(`Appearance: ${groupContentControl.appearance}`);
        console.log(`Cannot Delete: ${groupContentControl.cannotDelete}`);
    } else {
        console.log("No group content control found in the document");
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
  - `properties`: `Interfaces.GroupContentControlUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.GroupContentControl` (required)

  **Returns:** `void`

#### Examples

**Example**: Update multiple properties of a group content control, including its title and appearance settings

```typescript
await Word.run(async (context) => {
    // Get the first group content control in the document
    const groupContentControl = context.document.contentControls.getFirstOrNullObject();
    groupContentControl.load("type");
    
    await context.sync();
    
    if (groupContentControl.isNullObject) {
        console.log("No content control found");
        return;
    }
    
    // Set multiple properties at once using the set() method
    groupContentControl.set({
        title: "Employee Information",
        tag: "employee-group",
        appearance: Word.ContentControlAppearance.boundingBox,
        color: "blue"
    });
    
    await context.sync();
    console.log("Group content control properties updated");
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

**Example**: Set placeholder text "Enter your company name here" for a group content control to guide users on what information to provide.

```typescript
await Word.run(async (context) => {
    // Get the first group content control in the document
    const groupContentControl = context.document.contentControls.getFirstOrNullObject();
    groupContentControl.load("type");
    
    await context.sync();
    
    // Set placeholder text for the group content control
    if (groupContentControl.type === Word.ContentControlType.group) {
        groupContentControl.setPlaceholderText({
            placeholderText: "Enter your company name here"
        });
        
        await context.sync();
        console.log("Placeholder text set successfully");
    }
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.GroupContentControl object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.GroupContentControlData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.GroupContentControlData`

#### Examples

**Example**: Serialize a group content control to JSON format to log or store its properties

```typescript
await Word.run(async (context) => {
    // Get the first group content control in the document
    const groupContentControl = context.document.contentControls.getFirstOrNullObject();
    groupContentControl.load("id,tag,title,appearance,cannotDelete,cannotEdit");
    
    await context.sync();
    
    if (!groupContentControl.isNullObject) {
        // Convert the group content control to a plain JavaScript object
        const jsonData = groupContentControl.toJSON();
        
        // Now you can use the plain object (e.g., log it, store it, etc.)
        console.log("Group Content Control Data:", JSON.stringify(jsonData, null, 2));
        console.log("ID:", jsonData.id);
        console.log("Tag:", jsonData.tag);
    }
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.GroupContentControl`

#### Examples

**Example**: Track a group content control to maintain its reference across multiple sync calls while modifying its properties and content

```typescript
await Word.run(async (context) => {
    // Get the first group content control in the document
    const groupContentControl = context.document.contentControls.getFirstOrNullObject();
    groupContentControl.load("type");
    await context.sync();

    if (!groupContentControl.isNullObject && groupContentControl.type === "Group") {
        // Track the object to use it across multiple sync calls
        groupContentControl.track();

        // First sync - load properties
        groupContentControl.load("tag,title");
        await context.sync();

        console.log("Current tag:", groupContentControl.tag);

        // Second sync - modify properties
        groupContentControl.tag = "TrackedGroup";
        groupContentControl.title = "Updated Group";
        await context.sync();

        console.log("Updated tag:", groupContentControl.tag);

        // Untrack when done
        groupContentControl.untrack();
    }
});
```

---

### ungroup

Removes the group content control from the document so that its child content controls are no longer nested and can be freely edited.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Remove the grouping from the first group content control in the document to allow its child controls to be edited independently

```typescript
await Word.run(async (context) => {
    // Get the first group content control in the document
    const groupContentControls = context.document.contentControls.getByTypes([Word.ContentControlType.group]);
    const firstGroup = groupContentControls.getFirst();
    
    // Load the group content control
    firstGroup.load("id");
    await context.sync();
    
    // Ungroup the content control
    firstGroup.ungroup();
    
    await context.sync();
    console.log("Group content control has been ungrouped");
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.GroupContentControl`

#### Examples

**Example**: Track a group content control to work with it, then untrack it to release memory after modifications are complete.

```typescript
await Word.run(async (context) => {
    // Get the first group content control in the document
    const groupContentControl = context.document.contentControls.getFirstOrNullObject();
    groupContentControl.load("type");
    
    await context.sync();
    
    if (!groupContentControl.isNullObject && groupContentControl.type === "Group") {
        // Track the object to work with it across multiple sync calls
        context.trackedObjects.add(groupContentControl);
        
        // Perform operations with the group content control
        groupContentControl.load("tag");
        await context.sync();
        
        console.log("Group tag:", groupContentControl.tag);
        
        // Untrack the object to release memory when done
        groupContentControl.untrack();
        
        await context.sync();
    }
});
```

---

## Source

- /en-us/javascript/api/word
- /en-us/javascript/api/office/officeextension.clientobject
- /en-us/javascript/api/requirement-sets/word/word-api-requirement-sets
- /en-us/javascript/api/word/word.contentcontrolappearance
- /en-us/javascript/api/word/word.requestcontext
- /en-us/javascript/api/word/word.contentcontrollevel
- /en-us/javascript/api/word/word.buildingblock
- /en-us/javascript/api/word/word.range
- /en-us/javascript/api/word/word.xmlmapping
- /en-us/javascript/api/word/word.interfaces.groupcontentcontrolloadoptions
- /en-us/javascript/api/word/word.groupcontentcontrol
- /en-us/javascript/api/word/word.interfaces.groupcontentcontrolupdatedata
- /en-us/javascript/api/office/officeextension.updateoptions
- /en-us/javascript/api/word/word.contentcontrolplaceholderoptions
- /en-us/javascript/api/word/word.interfaces.groupcontentcontroldata
