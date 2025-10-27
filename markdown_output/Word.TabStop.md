# Word.TabStop

**Package:** `word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents a tab stop in a Word document.

## Properties

### alignment

**Type:** `Word.TabAlignment | "Left" | "Center" | "Right" | "Decimal" | "Bar" | "List"`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets a TabAlignment value that represents the alignment for the tab stop.

#### Examples

**Example**: Read and display the alignment type of the first tab stop in the first paragraph

```typescript
await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    const tabStops = firstParagraph.paragraphFormat.tabStops;
    tabStops.load("items");
    
    await context.sync();
    
    if (tabStops.items.length > 0) {
        const firstTabStop = tabStops.items[0];
        firstTabStop.load("alignment");
        
        await context.sync();
        
        console.log(`Tab stop alignment: ${firstTabStop.alignment}`);
        // Output example: "Tab stop alignment: Left" or "Center", "Right", etc.
    }
});
```

---

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a TabStop object to verify the connection between the add-in and Word application before performing operations on tab stops.

```typescript
await Word.run(async (context) => {
    // Get the first paragraph's tab stops
    const paragraph = context.document.body.paragraphs.getFirst();
    const tabStops = paragraph.tabStops;
    tabStops.load("items");
    
    await context.sync();
    
    if (tabStops.items.length > 0) {
        const firstTabStop = tabStops.items[0];
        
        // Access the request context from the TabStop object
        const tabStopContext = firstTabStop.context;
        
        // Verify the context is valid and connected
        if (tabStopContext) {
            console.log("TabStop is connected to Word application context");
            
            // Use the context to perform operations
            firstTabStop.load("position,alignment");
            await tabStopContext.sync();
            
            console.log(`Tab stop position: ${firstTabStop.position}`);
            console.log(`Tab stop alignment: ${firstTabStop.alignment}`);
        }
    }
});
```

---

### customTab

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets whether this tab stop is a custom tab stop.

#### Examples

**Example**: Check if the first tab stop in a paragraph is a custom tab stop and display the result in the console.

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    const tabStops = paragraph.tabStops;
    tabStops.load("items");
    
    await context.sync();
    
    if (tabStops.items.length > 0) {
        const firstTabStop = tabStops.items[0];
        firstTabStop.load("customTab");
        
        await context.sync();
        
        console.log(`Is custom tab stop: ${firstTabStop.customTab}`);
    } else {
        console.log("No tab stops found in the paragraph.");
    }
});
```

---

### leader

**Type:** `Word.TabLeader | "Spaces" | "Dots" | "Dashes" | "Lines" | "Heavy" | "MiddleDot"`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets a TabLeader value that represents the leader for this TabStop object.

#### Examples

**Example**: Read and display the leader style of the first tab stop in the selected paragraph

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.getSelection().paragraphs.getFirst();
    const tabStops = paragraph.tabStops;
    tabStops.load("items");
    
    await context.sync();
    
    if (tabStops.items.length > 0) {
        const firstTabStop = tabStops.items[0];
        firstTabStop.load("leader");
        
        await context.sync();
        
        console.log(`Tab stop leader style: ${firstTabStop.leader}`);
    } else {
        console.log("No tab stops found in the selected paragraph.");
    }
});
```

---

### next

**Type:** `Word.TabStop`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the next tab stop in the collection.

#### Examples

**Example**: Iterate through consecutive tab stops starting from the first one and log their positions to the console.

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    const tabStops = paragraph.tabStops;
    tabStops.load("items");
    
    await context.sync();
    
    if (tabStops.items.length > 0) {
        let currentTabStop = tabStops.items[0];
        currentTabStop.load("alignment, position");
        
        // Get the next tab stop
        let nextTabStop = currentTabStop.next;
        nextTabStop.load("alignment, position");
        
        await context.sync();
        
        console.log(`Current tab stop position: ${currentTabStop.position}`);
        console.log(`Next tab stop position: ${nextTabStop.position}`);
    }
});
```

---

### position

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the position of the tab stop relative to the left margin.

#### Examples

**Example**: Read and display the position of the first tab stop in the selected paragraph

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.getSelection().paragraphs.getFirst();
    const tabStops = paragraph.tabStops;
    tabStops.load("items");
    
    await context.sync();
    
    if (tabStops.items.length > 0) {
        const firstTabStop = tabStops.items[0];
        firstTabStop.load("position");
        
        await context.sync();
        
        console.log(`Tab stop position: ${firstTabStop.position} points from left margin`);
    } else {
        console.log("No tab stops found in the selected paragraph");
    }
});
```

---

### previous

**Type:** `Word.TabStop`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the previous tab stop in the collection.

#### Examples

**Example**: Navigate backwards through tab stops to find and remove the previous tab stop before a specific position

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    const tabStops = paragraph.tabStops;
    tabStops.load("items");
    
    await context.sync();
    
    // Get the last tab stop
    if (tabStops.items.length > 0) {
        const lastTabStop = tabStops.items[tabStops.items.length - 1];
        lastTabStop.load("previous");
        
        await context.sync();
        
        // Access the previous tab stop and delete it
        if (lastTabStop.previous) {
            lastTabStop.previous.delete();
            console.log("Previous tab stop deleted");
        }
        
        await context.sync();
    }
});
```

---

## Methods

### clear

**Kind:** `delete`

Removes this custom tab stop.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Remove the first custom tab stop from the first paragraph in the document

```typescript
await Word.run(async (context) => {
    // Get the first paragraph
    const paragraph = context.document.body.paragraphs.getFirst();
    
    // Get the tab stops collection
    const tabStops = paragraph.paragraphFormat.tabStops;
    tabStops.load("items");
    
    await context.sync();
    
    // Remove the first tab stop if it exists
    if (tabStops.items.length > 0) {
        tabStops.items[0].clear();
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
  - `options`: `Word.Interfaces.TabStopLoadOptions` (required)
    Provides options for which properties of the object to load.

  **Returns:** `Word.TabStop`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (required)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.TabStop`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (required)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.TabStop`

#### Examples

**Example**: Read and display the alignment and position properties of the first tab stop in the selected paragraph

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the selection
    const paragraph = context.document.getSelection().paragraphs.getFirst();
    
    // Get the first tab stop from the paragraph
    const tabStops = paragraph.tabStops;
    tabStops.load("items");
    await context.sync();
    
    if (tabStops.items.length > 0) {
        const firstTabStop = tabStops.items[0];
        
        // Load specific properties of the tab stop
        firstTabStop.load("alignment, position");
        await context.sync();
        
        // Display the loaded properties
        console.log(`Tab Stop Alignment: ${firstTabStop.alignment}`);
        console.log(`Tab Stop Position: ${firstTabStop.position}`);
    } else {
        console.log("No tab stops found in the selected paragraph.");
    }
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.TabStop object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.TabStopData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.TabStopData`

#### Examples

**Example**: Retrieve tab stop information from the first paragraph and convert it to a plain JavaScript object for logging or serialization purposes.

```typescript
await Word.run(async (context) => {
    // Get the first paragraph
    const paragraph = context.document.body.paragraphs.getFirst();
    
    // Get the tab stops collection
    const tabStops = paragraph.format.tabStops;
    tabStops.load("items");
    
    await context.sync();
    
    // Get the first tab stop if it exists
    if (tabStops.items.length > 0) {
        const tabStop = tabStops.items[0];
        tabStop.load("position,alignment,type");
        
        await context.sync();
        
        // Convert the TabStop object to a plain JavaScript object
        const tabStopData = tabStop.toJSON();
        
        // Now you can use the plain object for logging or serialization
        console.log("Tab Stop Data:", JSON.stringify(tabStopData, null, 2));
        console.log("Position:", tabStopData.position);
        console.log("Alignment:", tabStopData.alignment);
    }
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.TabStop`

#### Examples

**Example**: Track a tab stop object to maintain its reference across multiple sync calls when modifying paragraph formatting

```typescript
await Word.run(async (context) => {
    // Get the first paragraph
    const paragraph = context.document.body.paragraphs.getFirst();
    paragraph.load("tabStops");
    
    await context.sync();
    
    // Get the first tab stop (or add one if none exists)
    let tabStop: Word.TabStop;
    if (paragraph.tabStops.items.length > 0) {
        tabStop = paragraph.tabStops.items[0];
    } else {
        tabStop = paragraph.tabStops.add(144, Word.TabStopType.left);
    }
    
    // Track the tab stop to use it across multiple sync calls
    tabStop.track();
    
    await context.sync();
    
    // Now we can safely modify the tab stop in subsequent operations
    tabStop.alignment = Word.TabStopType.center;
    
    await context.sync();
    
    // Untrack when done
    tabStop.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.TabStop`

#### Examples

**Example**: Add a custom tab stop to a paragraph, use it, and then untrack the tab stop object to free memory after the operation is complete.

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    paragraph.load("tabStops");
    
    await context.sync();
    
    // Add and track a tab stop at 144 points (2 inches)
    const tabStop = paragraph.tabStops.add(144, Word.TabStopType.left);
    
    // Insert text with a tab character to use the tab stop
    paragraph.insertText("\tTabbed text", Word.InsertLocation.end);
    
    await context.sync();
    
    // Untrack the tab stop object to release memory
    tabStop.untrack();
    
    await context.sync();
});
```

---

## Source

- /en-us/javascript/api/word
