# Word.StyleCollection

**Package:** `word`

**API Set:** WordApi 1.5 None

**Extends:** `OfficeExtension.ClientObject`

## Description

Contains a collection of [Word.Style](https://learn.microsoft.com/en-us/javascript/api/word/word.style) objects.

## Class Examples

```typescript
// Link to full sample: // Gets the number of available styles stored with the document.
await Word.run(async (context) => {
  const styles: Word.StyleCollection = context.document.getStyles();
  const count = styles.getCount();
  await context.sync();

  console.log(`Number of styles: ${count.value}`);
});
```

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a StyleCollection to verify the connection between the add-in and Word application is active before performing style operations

```typescript
await Word.run(async (context) => {
    const styles = context.document.getStyles();
    
    // Access the request context associated with the StyleCollection
    const requestContext = styles.context;
    
    // Verify the context is valid and connected
    if (requestContext && requestContext.application) {
        console.log("StyleCollection is connected to Word application");
        
        // Load styles using the same context
        styles.load("items");
        await context.sync();
        
        console.log(`Successfully accessed ${styles.items.length} styles`);
    }
});
```

---

### items

**Type:** `Word.Style[]`

Gets the loaded child items in this collection.

#### Examples

**Example**: List all available style names in the document to the console

```typescript
await Word.run(async (context) => {
    const styles = context.document.getStyles();
    styles.load("items");
    
    await context.sync();
    
    console.log("Available styles in the document:");
    styles.items.forEach((style) => {
        console.log(`- ${style.nameLocal}`);
    });
});
```

---

## Methods

### getByName

**Kind:** `read`

Get the style object by its name.

#### Signature

**Parameters:**
- `name`: `string` (required)
  The style name.

**Returns:** `Word.Style`

#### Examples

**Example**: Get the "Heading 1" style from the document's style collection and apply it to the first paragraph.

```typescript
await Word.run(async (context) => {
    // Get the style collection
    const styles = context.document.getStyles();
    
    // Get the "Heading 1" style by name
    const heading1Style = styles.getByName("Heading 1");
    
    // Get the first paragraph in the document
    const firstParagraph = context.document.body.paragraphs.getFirst();
    
    // Load the style name to verify
    heading1Style.load("nameLocal");
    
    await context.sync();
    
    // Apply the style to the first paragraph
    firstParagraph.style = heading1Style.nameLocal;
    
    await context.sync();
});
```

---

### getByNameOrNullObject

**Kind:** `read`

If the corresponding style doesn't exist, then this method returns an object with its isNullObject property set to true.

#### Signature

**Parameters:**
- `name`: `string` (required)
  The style name.

**Returns:** `Word.Style`

#### Examples

**Example**: Check if a style with the given name already exists in the document, and if not, add a new style with the specified name and type.

```typescript
// Link to full sample: // Adds a new style.
await Word.run(async (context) => {
  const newStyleName = (document.getElementById("new-style-name") as HTMLInputElement).value;
  if (newStyleName == "") {
    console.warn("Enter a style name to add.");
    return;
  }

  const style: Word.Style = context.document.getStyles().getByNameOrNullObject(newStyleName);
  style.load();
  await context.sync();

  if (!style.isNullObject) {
    console.warn(
      `There's an existing style with the same name '${newStyleName}'! Please provide another style name.`
    );
    return;
  }

  const newStyleType = ((document.getElementById("new-style-type") as HTMLSelectElement).value as unknown) as Word.StyleType;
  context.document.addStyle(newStyleName, newStyleType);
  await context.sync();

  console.log(newStyleName + " has been added to the style list.");
});
```

---

### getCount

**Kind:** `read`

Gets the number of the styles in the collection.

#### Signature

**Returns:** `OfficeExtension.ClientResult<number>`

#### Examples

**Example**: Retrieve and display the total number of styles available in the document.

```typescript
// Link to full sample: // Gets the number of available styles stored with the document.
await Word.run(async (context) => {
  const styles: Word.StyleCollection = context.document.getStyles();
  const count = styles.getCount();
  await context.sync();

  console.log(`Number of styles: ${count.value}`);
});
```

---

### getItem

**Kind:** `read`

Gets a style object by its index in the collection.

#### Signature

**Parameters:**
- `index`: `number` (required)
  A number that identifies the index location of a style object.

**Returns:** `Word.Style`

#### Examples

**Example**: Get the third style in the document's style collection and apply it to the first paragraph

```typescript
await Word.run(async (context) => {
    // Get the style collection
    const styles = context.document.getStyles();
    
    // Get the style at index 2 (third style)
    const thirdStyle = styles.getItem(2);
    thirdStyle.load("nameLocal");
    
    // Get the first paragraph
    const firstParagraph = context.document.body.paragraphs.getFirst();
    
    await context.sync();
    
    // Apply the style to the paragraph
    firstParagraph.style = thirdStyle.nameLocal;
    
    await context.sync();
    
    console.log(`Applied style: ${thirdStyle.nameLocal}`);
});
```

---

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.StyleCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.StyleCollection`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.StyleCollection`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `OfficeExtension.LoadOption` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.StyleCollection`

#### Examples

**Example**: Load and display the names of all available styles in the document

```typescript
await Word.run(async (context) => {
    // Get the style collection
    const styles = context.document.getStyles();
    
    // Load the name property for all styles in the collection
    styles.load("items/name");
    
    // Synchronize the document state
    await context.sync();
    
    // Display the style names
    console.log("Available styles:");
    styles.items.forEach(style => {
        console.log(style.name);
    });
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.StyleCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.StyleCollectionData) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

#### Signature

**Returns:** `Word.Interfaces.StyleCollectionData`

#### Examples

**Example**: Export all available document styles to JSON format for logging or external storage

```typescript
await Word.run(async (context) => {
    // Get the style collection from the document
    const styles = context.document.getStyles();
    
    // Load the properties we want to export
    styles.load("items/nameLocal,items/type,items/builtIn");
    
    await context.sync();
    
    // Convert the style collection to a plain JavaScript object
    const stylesJSON = styles.toJSON();
    
    // Log or use the JSON data
    console.log("Document Styles:", JSON.stringify(stylesJSON, null, 2));
    
    // Example: Access the items array from the JSON object
    console.log(`Total styles: ${stylesJSON.items.length}`);
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.StyleCollection`

#### Examples

**Example**: Track a StyleCollection object to safely access and modify styles across multiple sync calls, preventing "InvalidObjectPath" errors when working with the collection outside a single batch operation.

```typescript
await Word.run(async (context) => {
    const styles = context.document.getStyles();
    
    // Track the StyleCollection to use it across multiple sync calls
    styles.track();
    
    // Load the collection
    styles.load("items");
    await context.sync();
    
    // First sync - log the count
    console.log(`Total styles: ${styles.items.length}`);
    
    // Perform another operation after sync
    await context.sync();
    
    // Can still safely access the tracked collection
    for (const style of styles.items) {
        style.load("nameLocal,type");
    }
    await context.sync();
    
    // Access properties after multiple syncs
    styles.items.forEach(style => {
        console.log(`Style: ${style.nameLocal}, Type: ${style.type}`);
    });
    
    // Untrack when done to free up memory
    styles.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.StyleCollection`

#### Examples

**Example**: Load and use the styles collection, then release it from memory tracking to improve performance

```typescript
await Word.run(async (context) => {
    // Load the styles collection and track it
    const styles = context.document.getStyles();
    styles.load("items");
    
    await context.sync();
    
    // Use the styles collection
    console.log(`Total styles: ${styles.items.length}`);
    
    // Release the memory associated with the styles collection
    styles.untrack();
    
    await context.sync();
});
```

---

## Source

- https://learn.microsoft.com/en-us/javascript/api/word/word.stylecollection
