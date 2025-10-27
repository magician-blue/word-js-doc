# Word.BorderUniversalCollection

**Package:** `word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents the collection of [Word.BorderUniversal](/en-us/javascript/api/word/word.borderuniversal) objects.

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a BorderUniversalCollection to verify the connection to the Word host application before performing border operations.

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    const borders = paragraph.getBorderUniversalCollection();
    
    // Access the request context associated with the border collection
    const borderContext = borders.context;
    
    // Verify the context is valid and connected to the host application
    console.log("Border collection context is connected:", borderContext !== null);
    
    // Use the context to load and sync border properties
    borders.load("items");
    await borderContext.sync();
    
    console.log(`Found ${borders.items.length} borders in the collection`);
});
```

---

### items

**Type:** `Word.BorderUniversal[]`

Gets the loaded child items in this collection.

#### Examples

**Example**: Iterate through all borders in a paragraph's border collection and set each border's color to blue

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    const borders = paragraph.getBorders();
    borders.load("items");
    
    await context.sync();
    
    // Access the loaded border items and set their color
    for (const border of borders.items) {
        border.color = "blue";
    }
    
    await context.sync();
});
```

---

## Methods

### applyPageBordersToAllSections

Applies the specified page-border formatting to all sections in the document.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Apply a blue double-line page border with 24pt width to all sections in the document

```typescript
await Word.run(async (context) => {
    // Get the border collection
    const borders = context.document.body.parentContentControlOrNullObject.parentBody.parentContentControlOrNullObject.parentBody.borders;
    
    // Configure the page border settings
    borders.load("items");
    await context.sync();
    
    // Set border properties (e.g., for top border)
    borders.items[0].type = Word.BorderType.double;
    borders.items[0].color = "#0000FF"; // Blue
    borders.items[0].width = 24;
    
    // Apply the page border settings to all sections
    borders.applyPageBordersToAllSections();
    
    await context.sync();
});
```

---

### getItem

**Kind:** `read`

Gets a `Border` object by its index in the collection.

#### Signature

**Parameters:**
- `index`: `number` (required)
  The location of a `BorderUniversal` object.

**Returns:** `Word.BorderUniversal`

#### Examples

**Example**: Get the top border of the first paragraph and set its color to red

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    const borders = paragraph.getBorders();
    
    // Get the top border (index 0) from the collection
    const topBorder = borders.getItem(0);
    topBorder.color = "red";
    
    await context.sync();
});
```

---

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.BorderUniversalCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.BorderUniversalCollection`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.BorderUniversalCollection`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `OfficeExtension.LoadOption` (optional)
    `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.BorderUniversalCollection`

#### Examples

**Example**: Load and display the border widths of all borders in the first paragraph

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    const borders = paragraph.borders;
    
    // Load the width property for all borders in the collection
    borders.load("items/width");
    
    await context.sync();
    
    // Display the border widths
    borders.items.forEach((border, index) => {
        console.log(`Border ${index} width: ${border.width}`);
    });
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.BorderUniversalCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.BorderUniversalCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

#### Signature

**Returns:** `Word.Interfaces.BorderUniversalCollectionData`

#### Examples

**Example**: Serialize a paragraph's border collection to JSON format for logging or data export purposes.

```typescript
await Word.run(async (context) => {
    // Get the first paragraph
    const paragraph = context.document.body.paragraphs.getFirst();
    
    // Get the border collection
    const borders = paragraph.borders;
    
    // Load the border properties
    borders.load("items");
    
    await context.sync();
    
    // Convert the border collection to a plain JavaScript object
    const bordersJSON = borders.toJSON();
    
    // Log or export the serialized data
    console.log(JSON.stringify(bordersJSON, null, 2));
    
    // You can now work with the plain object
    console.log(`Number of borders: ${bordersJSON.items.length}`);
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.BorderUniversalCollection`

#### Examples

**Example**: Track a paragraph's border collection to maintain references across multiple sync calls when applying different border styles sequentially

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    const borders = paragraph.getBorder().borders;
    
    // Track the border collection to use it across multiple sync calls
    borders.track();
    
    await context.sync();
    
    // Now we can safely use the borders object across sync calls
    borders.items[0].color = "blue";
    await context.sync();
    
    borders.items[1].width = 2;
    await context.sync();
    
    // Untrack when done to free up memory
    borders.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

#### Signature

**Returns:** `Word.BorderUniversalCollection`

#### Examples

**Example**: Get all borders from a table, track the collection for performance monitoring, then release the tracking when done to free memory.

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    
    // Get the borders collection and track it
    const borders = table.getBorder().load("items");
    context.trackedObjects.add(borders);
    
    await context.sync();
    
    // Work with the borders collection
    console.log(`Found ${borders.items.length} borders`);
    
    // Release tracking when done to free memory
    borders.untrack();
    
    await context.sync();
});
```

---

## Source

- /en-us/javascript/api/word
