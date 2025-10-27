# TabStopCollection

**Package:** `word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents a collection of [tab stops](/en-us/javascript/api/word/word.tabstop) in a Word document.

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a TabStopCollection to verify the connection between the add-in and Word, then use it to load and log tab stop information.

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    const tabStops = paragraph.tabStops;
    
    // Access the request context associated with the TabStopCollection
    const tabStopContext = tabStops.context;
    
    // Use the context to load properties
    tabStops.load("items");
    await tabStopContext.sync();
    
    // Log the number of tab stops using the context connection
    console.log(`Number of tab stops: ${tabStops.items.length}`);
    console.log(`Context is connected: ${tabStopContext !== null}`);
});
```

---

### items

**Type:** `Word.TabStop[]`

Gets the loaded child items in this collection.

#### Examples

**Example**: Iterate through all tab stops in the first paragraph and log their position and alignment type to the console.

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    const tabStops = paragraph.tabStops;
    
    // Load the tab stops collection
    tabStops.load("items");
    await context.sync();
    
    // Access the loaded tab stops using the items property
    const tabStopItems = tabStops.items;
    
    console.log(`Found ${tabStopItems.length} tab stops`);
    
    for (let i = 0; i < tabStopItems.length; i++) {
        const tabStop = tabStopItems[i];
        tabStop.load("position, alignment");
    }
    
    await context.sync();
    
    // Log details of each tab stop
    tabStopItems.forEach((tabStop, index) => {
        console.log(`Tab Stop ${index + 1}: Position = ${tabStop.position}, Alignment = ${tabStop.alignment}`);
    });
});
```

---

## Methods

### add

**Kind:** `create`

Returns a `TabStop` object that represents a custom tab stop added to the paragraph.

#### Signature

**Parameters:**
- `position`: `number` (required)
  The position of the tab stop.
- `options`: `Word.TabStopAddOptions` (optional)
  Optional. The options to further configure the new tab stop.

**Returns:** `Word.TabStop`

#### Examples

**Example**: Add a custom tab stop at 3 inches (216 points) to the first paragraph with left alignment

```typescript
await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    
    // Add a tab stop at 3 inches (216 points) with left alignment
    firstParagraph.tabStops.add(216, { alignment: Word.Alignment.left });
    
    await context.sync();
    console.log("Tab stop added successfully");
});
```

---

### after

**Kind:** `read`

Returns the next `TabStop` object to the right of the specified position.

#### Signature

**Parameters:**
- `Position`: `number` (required)
  The position to check.

**Returns:** `Word.TabStop`

#### Examples

**Example**: Find and highlight the next tab stop that appears after the 2-inch position in the first paragraph

```typescript
await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    const tabStops = firstParagraph.paragraphFormat.tabStops;
    
    // Get the next tab stop after the 2-inch position (144 points)
    const nextTabStop = tabStops.after(144);
    
    // Load the position property to verify
    nextTabStop.load("position");
    
    await context.sync();
    
    console.log(`Next tab stop after 2 inches is at: ${nextTabStop.position} points`);
});
```

---

### before

**Kind:** `read`

Returns the next `TabStop` object to the left of the specified position.

#### Signature

**Parameters:**
- `Position`: `number` (required)
  The position to check.

**Returns:** `Word.TabStop`

#### Examples

**Example**: Find and highlight the tab stop that appears immediately before the 3-inch position in the first paragraph

```typescript
await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    const tabStops = firstParagraph.paragraphFormat.tabStops;
    
    // Get the tab stop before the 3-inch position (216 points)
    const tabStopBefore = tabStops.before(216);
    
    tabStopBefore.load("position,alignment");
    await context.sync();
    
    console.log(`Found tab stop at ${tabStopBefore.position} points with ${tabStopBefore.alignment} alignment`);
});
```

---

### clearAll

**Kind:** `delete`

Clears all the custom tab stops from the paragraph.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Remove all custom tab stops from the first paragraph in the document to reset its tab formatting to default settings.

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const paragraph = context.document.body.paragraphs.getFirst();
    
    // Get the tab stop collection for the paragraph
    const tabStops = paragraph.tabStops;
    
    // Clear all custom tab stops
    tabStops.clearAll();
    
    await context.sync();
    
    console.log("All custom tab stops have been cleared from the paragraph.");
});
```

---

### getItem

**Kind:** `read`

Gets a `TabStop` object by its index in the collection.

#### Signature

**Parameters:**
- `index`: `number` (required)
  A number that identifies the index location of a `TabStop` object.

**Returns:** `Word.TabStop`

#### Examples

**Example**: Get the second tab stop from the first paragraph and change its alignment to center

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const paragraph = context.document.body.paragraphs.getFirst();
    
    // Get the tab stops collection
    const tabStops = paragraph.tabStops;
    
    // Get the second tab stop (index 1)
    const secondTabStop = tabStops.getItem(1);
    
    // Change its alignment to center
    secondTabStop.alignment = Word.Alignment.center;
    
    // Load properties to verify
    secondTabStop.load("position,alignment");
    
    await context.sync();
    
    console.log(`Tab stop at position ${secondTabStop.position} is now ${secondTabStop.alignment}`);
});
```

---

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.TabStopCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.TabStopCollection`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.TabStopCollection`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `OfficeExtension.LoadOption` (optional)
    `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.TabStopCollection`

#### Examples

**Example**: Load and display the position and alignment of all tab stops in the first paragraph

```typescript
await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    const tabStops = firstParagraph.tabStops;
    
    // Load properties of the tab stop collection
    tabStops.load("items/position, items/alignment");
    
    await context.sync();
    
    // Display the tab stop information
    console.log(`Found ${tabStops.items.length} tab stops:`);
    tabStops.items.forEach((tabStop, index) => {
        console.log(`Tab Stop ${index + 1}: Position = ${tabStop.position}, Alignment = ${tabStop.alignment}`);
    });
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.TabStopCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.TabStopCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

#### Signature

**Returns:** `Word.Interfaces.TabStopCollectionData`

#### Examples

**Example**: Serialize tab stop collection data to JSON format for logging or external storage

```typescript
await Word.run(async (context) => {
    // Get the first paragraph
    const paragraph = context.document.body.paragraphs.getFirst();
    
    // Get the tab stops collection
    const tabStops = paragraph.paragraphFormat.tabStops;
    
    // Load the tab stops properties
    tabStops.load("items");
    
    await context.sync();
    
    // Convert the tab stops collection to a plain JavaScript object
    const tabStopsData = tabStops.toJSON();
    
    // Log the serialized data
    console.log("Tab Stops Data:", JSON.stringify(tabStopsData, null, 2));
    
    // The tabStopsData object contains an "items" array with tab stop properties
    console.log(`Number of tab stops: ${tabStopsData.items.length}`);
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.TabStopCollection`

#### Examples

**Example**: Track tab stops in a paragraph across multiple sync calls to safely modify them without getting InvalidObjectPath errors

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    const tabStops = paragraph.tabStops;
    
    // Track the tab stops collection to use it across multiple sync calls
    tabStops.track();
    
    await context.sync();
    
    // Now we can safely work with the tracked object across sync calls
    tabStops.add(144, Word.Alignment.left); // Add tab at 2 inches
    
    await context.sync();
    
    // Still safe to use the tabStops object
    tabStops.add(288, Word.Alignment.center); // Add tab at 4 inches
    
    await context.sync();
    
    // Clean up tracking when done
    tabStops.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

#### Signature

**Returns:** `Word.TabStopCollection`

#### Examples

**Example**: Load tab stops from a paragraph, use them to display information, then untrack the collection to free memory

```typescript
await Word.run(async (context) => {
    // Get the first paragraph
    const paragraph = context.document.body.paragraphs.getFirst();
    const tabStops = paragraph.tabStops;
    
    // Load the tab stops collection
    tabStops.load("items");
    await context.sync();
    
    // Use the tab stops data
    console.log(`Found ${tabStops.items.length} tab stops`);
    
    // Untrack the collection to release memory
    tabStops.untrack();
    await context.sync();
});
```

---

## Source

- /en-us/javascript/api/word
