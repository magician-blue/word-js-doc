# Hyperlink

**Package:** `word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents a hyperlink in a Word document.

## Properties

### address

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the address (for example, a file name or URL) of the hyperlink.

#### Examples

**Example**: Update an existing hyperlink's address to point to a different URL

```typescript
await Word.run(async (context) => {
    // Get the first hyperlink in the document
    const hyperlinks = context.document.body.getHyperlinks();
    const firstHyperlink = hyperlinks.getFirst();
    
    // Update the hyperlink address to a new URL
    firstHyperlink.address = "https://www.microsoft.com";
    
    await context.sync();
});
```

---

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the hyperlink's request context to verify the connection between the add-in and Word before performing operations on the hyperlink.

```typescript
await Word.run(async (context) => {
    // Get the first hyperlink in the document
    const hyperlinks = context.document.body.getHyperlinks();
    hyperlinks.load("items");
    await context.sync();
    
    if (hyperlinks.items.length > 0) {
        const hyperlink = hyperlinks.items[0];
        
        // Access the request context associated with the hyperlink
        const hyperlinkContext = hyperlink.context;
        
        // Verify the context is valid and connected
        console.log("Context is connected:", hyperlinkContext !== null);
        
        // Use the context to load properties
        hyperlink.load("address, displayText");
        await hyperlinkContext.sync();
        
        console.log("Hyperlink address:", hyperlink.address);
        console.log("Hyperlink text:", hyperlink.displayText);
    }
});
```

---

### emailSubject

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the text string for the hyperlink's subject line.

#### Examples

**Example**: Set the subject line of an email hyperlink to "Quarterly Report Review"

```typescript
await Word.run(async (context) => {
    // Get the first hyperlink in the document
    const hyperlinks = context.document.body.getHyperlinks();
    const firstHyperlink = hyperlinks.getFirst();
    
    // Set the email subject line
    firstHyperlink.emailSubject = "Quarterly Report Review";
    
    await context.sync();
});
```

---

### isExtraInfoRequired

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns true if extra information is required to resolve the hyperlink.

#### Examples

**Example**: Check if a hyperlink requires extra information to be resolved and display an alert to the user if additional details are needed.

```typescript
await Word.run(async (context) => {
    // Get the first hyperlink in the document
    const hyperlinks = context.document.body.getHyperlinks();
    const firstHyperlink = hyperlinks.getFirst();
    
    // Load the isExtraInfoRequired property
    firstHyperlink.load("isExtraInfoRequired");
    
    await context.sync();
    
    // Check if extra information is required
    if (firstHyperlink.isExtraInfoRequired) {
        console.log("This hyperlink requires additional information to resolve.");
    } else {
        console.log("This hyperlink can be resolved without extra information.");
    }
});
```

---

### name

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns the name of the Hyperlink object.

#### Examples

**Example**: Get the name of the first hyperlink in the document and display it in the console

```typescript
await Word.run(async (context) => {
    const hyperlinks = context.document.body.hyperlinks;
    hyperlinks.load("items");
    
    await context.sync();
    
    if (hyperlinks.items.length > 0) {
        const firstHyperlink = hyperlinks.items[0];
        firstHyperlink.load("name");
        
        await context.sync();
        
        console.log("Hyperlink name: " + firstHyperlink.name);
    }
});
```

---

### range

**Type:** `Word.Range`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns a Range object that represents the portion of the document that's contained within the hyperlink.

#### Examples

**Example**: Get the text content of a hyperlink by accessing its range property and then highlight that range in yellow.

```typescript
await Word.run(async (context) => {
    // Get the first hyperlink in the document
    const hyperlinks = context.document.body.hyperlinks;
    const firstHyperlink = hyperlinks.getFirst();
    
    // Get the range of the hyperlink
    const hyperlinkRange = firstHyperlink.range;
    
    // Highlight the hyperlink range in yellow
    hyperlinkRange.font.highlightColor = "yellow";
    
    // Load the text to display it
    hyperlinkRange.load("text");
    
    await context.sync();
    
    console.log("Hyperlink text: " + hyperlinkRange.text);
});
```

---

### screenTip

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the text that appears as a ScreenTip when the mouse pointer is positioned over the hyperlink.

#### Examples

**Example**: Set a custom tooltip message "Click to visit our documentation" for a hyperlink's ScreenTip

```typescript
await Word.run(async (context) => {
    // Get the first hyperlink in the document
    const hyperlinks = context.document.body.getHyperlinks();
    const firstHyperlink = hyperlinks.getFirst();
    
    // Set the ScreenTip text
    firstHyperlink.screenTip = "Click to visit our documentation";
    
    await context.sync();
});
```

---

### subAddress

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies a named location in the destination of the hyperlink.

#### Examples

**Example**: Set a hyperlink to jump to a specific bookmark named "Section3" within the same document

```typescript
await Word.run(async (context) => {
    // Get the first hyperlink in the document
    const hyperlink = context.document.body.getHyperlinks().getFirst();
    
    // Set the subAddress to point to a bookmark named "Section3"
    hyperlink.subAddress = "Section3";
    
    await context.sync();
    
    console.log("Hyperlink subAddress set to Section3");
});
```

---

### target

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the name of the frame or window in which to load the hyperlink.

#### Examples

**Example**: Set a hyperlink to open in a new browser window named "_blank"

```typescript
await Word.run(async (context) => {
    // Get the first hyperlink in the document
    const hyperlinks = context.document.body.getHyperlinks();
    const firstHyperlink = hyperlinks.getFirst();
    
    // Set the target to open in a new window
    firstHyperlink.target = "_blank";
    
    await context.sync();
});
```

---

### textToDisplay

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the hyperlink's visible text in the document.

#### Examples

**Example**: Change the display text of the first hyperlink in the document to "Click here for more information"

```typescript
await Word.run(async (context) => {
    const hyperlinks = context.document.body.getHyperlinks();
    hyperlinks.load("items");
    
    await context.sync();
    
    if (hyperlinks.items.length > 0) {
        const firstHyperlink = hyperlinks.items[0];
        firstHyperlink.textToDisplay = "Click here for more information";
        
        await context.sync();
    }
});
```

---

### type

**Type:** `Word.HyperlinkType | "Range" | "Shape" | "InlineShape"`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns the hyperlink type.

#### Examples

**Example**: Check the type of a hyperlink in the document and display different messages based on whether it's a Range, Shape, or InlineShape hyperlink.

```typescript
await Word.run(async (context) => {
    const hyperlinks = context.document.body.hyperlinks;
    hyperlinks.load("items");
    
    await context.sync();
    
    if (hyperlinks.items.length > 0) {
        const firstHyperlink = hyperlinks.items[0];
        firstHyperlink.load("type");
        
        await context.sync();
        
        const hyperlinkType = firstHyperlink.type;
        
        if (hyperlinkType === Word.HyperlinkType.range || hyperlinkType === "Range") {
            console.log("This is a text-based hyperlink");
        } else if (hyperlinkType === Word.HyperlinkType.shape || hyperlinkType === "Shape") {
            console.log("This is a shape hyperlink");
        } else if (hyperlinkType === Word.HyperlinkType.inlineShape || hyperlinkType === "InlineShape") {
            console.log("This is an inline shape hyperlink");
        }
    }
});
```

---

## Methods

### addToFavorites

Creates a shortcut to the document or hyperlink and adds it to the Favorites folder.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Add the first hyperlink found in the document to the user's Favorites folder

```typescript
await Word.run(async (context) => {
    // Get the first hyperlink in the document
    const hyperlinks = context.document.body.hyperlinks;
    hyperlinks.load("items");
    
    await context.sync();
    
    if (hyperlinks.items.length > 0) {
        const firstHyperlink = hyperlinks.items[0];
        
        // Add the hyperlink to Favorites
        firstHyperlink.addToFavorites();
        
        await context.sync();
        console.log("Hyperlink added to Favorites successfully");
    } else {
        console.log("No hyperlinks found in the document");
    }
});
```

---

### createNewDocument

**Kind:** `create`

Creates a new document linked to the hyperlink.

#### Signature

**Parameters:**
- `fileName`: `string` (required)
  The name of the file.
- `editNow`: `boolean` (required)
  true to start editing now.
- `overwrite`: `boolean` (required)
  true to overwrite if there's another file with the same name.

**Returns:** `void`

#### Examples

**Example**: Create a hyperlink in the document that, when clicked, creates a new Word document named "ProjectNotes.docx" and opens it for editing

```typescript
await Word.run(async (context) => {
    // Get the first paragraph and insert a hyperlink
    const paragraph = context.document.body.paragraphs.getFirst();
    const hyperlink = paragraph.insertText("Click to create new document", Word.InsertLocation.end)
        .insertHyperlink("", Word.InsertLocation.replace, "Create Project Notes");
    
    // Create a new document linked to this hyperlink
    // fileName: "ProjectNotes.docx" - name of the new document
    // editNow: true - open the document immediately for editing
    // overwrite: false - don't overwrite if file already exists
    hyperlink.createNewDocument("ProjectNotes.docx", true, false);
    
    await context.sync();
});
```

---

### delete

**Kind:** `delete`

Deletes the hyperlink.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Delete all hyperlinks in the document that contain "example.com" in their address

```typescript
await Word.run(async (context) => {
    // Get all hyperlinks in the document
    const hyperlinks = context.document.body.getHyperlinks();
    hyperlinks.load("items");
    
    await context.sync();
    
    // Delete hyperlinks that contain "example.com"
    for (let i = 0; i < hyperlinks.items.length; i++) {
        const hyperlink = hyperlinks.items[i];
        hyperlink.load("address");
        await context.sync();
        
        if (hyperlink.address.includes("example.com")) {
            hyperlink.delete();
        }
    }
    
    await context.sync();
});
```

---

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.HyperlinkLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.Hyperlink`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.Hyperlink`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `object` (optional)
    select is a comma-delimited string that specifies the properties to load, and expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.Hyperlink`

#### Examples

**Example**: Load and display the URL and text of the first hyperlink in the document

```typescript
await Word.run(async (context) => {
    // Get the first hyperlink in the document
    const hyperlinks = context.document.body.hyperlinks;
    hyperlinks.load("items");
    await context.sync();
    
    if (hyperlinks.items.length > 0) {
        const firstHyperlink = hyperlinks.items[0];
        
        // Load specific properties of the hyperlink
        firstHyperlink.load("address, textToDisplay");
        await context.sync();
        
        // Display the loaded properties
        console.log("URL: " + firstHyperlink.address);
        console.log("Display Text: " + firstHyperlink.textToDisplay);
    } else {
        console.log("No hyperlinks found in the document");
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
  - `properties`: `Interfaces.HyperlinkUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.Hyperlink` (required)

  **Returns:** `void`

#### Examples

**Example**: Update an existing hyperlink's display text and screen tip properties

```typescript
await Word.run(async (context) => {
    // Get the first hyperlink in the document
    const hyperlinks = context.document.body.getHyperlinks();
    const firstHyperlink = hyperlinks.getFirst();
    
    // Set multiple properties at once using the set() method
    firstHyperlink.set({
        textToDisplay: "Visit Microsoft",
        screenTip: "Click to visit Microsoft's official website"
    });
    
    await context.sync();
    console.log("Hyperlink properties updated successfully");
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.Hyperlink object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.HyperlinkData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.HyperlinkData`

#### Examples

**Example**: Serialize a hyperlink's properties to a plain JavaScript object and log it to the console for debugging purposes.

```typescript
await Word.run(async (context) => {
    // Get the first hyperlink in the document
    const hyperlinks = context.document.body.getHyperlinks();
    const firstHyperlink = hyperlinks.getFirst();
    
    // Load properties we want to serialize
    firstHyperlink.load("address, screenTip, textToDisplay");
    
    await context.sync();
    
    // Convert the hyperlink to a plain JavaScript object
    const hyperlinkData = firstHyperlink.toJSON();
    
    // Now we can safely log or manipulate the plain object
    console.log("Hyperlink data:", hyperlinkData);
    console.log("Address:", hyperlinkData.address);
    console.log("Screen tip:", hyperlinkData.screenTip);
    console.log("Display text:", hyperlinkData.textToDisplay);
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.Hyperlink`

#### Examples

**Example**: Track a hyperlink object across multiple sync calls to modify its properties without encountering InvalidObjectPath errors

```typescript
await Word.run(async (context) => {
    // Get the first hyperlink in the document
    const hyperlinks = context.document.body.hyperlinks;
    hyperlinks.load("items");
    await context.sync();
    
    if (hyperlinks.items.length > 0) {
        const hyperlink = hyperlinks.items[0];
        
        // Track the hyperlink to use it across multiple sync calls
        hyperlink.track();
        
        // Load properties
        hyperlink.load("textToDisplay, address");
        await context.sync();
        
        console.log("Original text: " + hyperlink.textToDisplay);
        console.log("Original address: " + hyperlink.address);
        
        // Modify the hyperlink after another sync
        hyperlink.textToDisplay = "Updated Link Text";
        await context.sync();
        
        console.log("Updated text: " + hyperlink.textToDisplay);
        
        // Untrack when done
        hyperlink.untrack();
    }
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.Hyperlink`

#### Examples

**Example**: Load hyperlink properties, use them, then untrack the hyperlink object to free memory

```typescript
await Word.run(async (context) => {
    // Get the first hyperlink in the document
    const hyperlinks = context.document.body.getHyperlinks();
    const firstHyperlink = hyperlinks.getFirst();
    
    // Track the object to work with it
    firstHyperlink.track();
    firstHyperlink.load("address, textToDisplay");
    
    await context.sync();
    
    // Use the hyperlink data
    console.log(`Link text: ${firstHyperlink.textToDisplay}`);
    console.log(`Link address: ${firstHyperlink.address}`);
    
    // Untrack to release memory after we're done using it
    firstHyperlink.untrack();
    
    await context.sync();
});
```

---

## Source

- https://learn.microsoft.com/en-us/javascript/api/word
- https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientobject
- https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets
- https://learn.microsoft.com/en-us/javascript/api/word/word.requestcontext
- https://learn.microsoft.com/en-us/javascript/api/word/word.range
- https://learn.microsoft.com/en-us/javascript/api/word/word.hyperlinktype
- https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.hyperlinkloadoptions
- https://learn.microsoft.com/en-us/javascript/api/word/word.hyperlink
- https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.hyperlinkupdatedata
- https://learn.microsoft.com/en-us/javascript/api/office/officeextension.updateoptions
- https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.hyperlinkdata
- https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member
