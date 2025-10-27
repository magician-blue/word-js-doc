# Word.ParagraphCollection

**Package:** `word`

**API Set:** WordApi 1.1 None

**Extends:** `OfficeExtension.ClientObject`

## Description

Contains a collection of [Word.Paragraph](/en-us/javascript/api/word/word.paragraph) objects.

## Class Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/15-images/insert-and-get-pictures.yaml

// Inserts an image anchored to the last paragraph.
await Word.run(async (context) => {
  context.document.body.paragraphs
    .getLast()
    .insertParagraph("", "After")
    .insertInlinePictureFromBase64(base64Image, "End");

  await context.sync();
});
```

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a ParagraphCollection to verify the connection between the add-in and Word, then use it to load and sync paragraph data.

```typescript
await Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    
    // Access the request context associated with the ParagraphCollection
    const requestContext = paragraphs.context;
    
    // Use the context to load properties
    paragraphs.load("text");
    
    // Sync using the request context
    await requestContext.sync();
    
    // Log the first paragraph to verify the context connection worked
    if (paragraphs.items.length > 0) {
        console.log("First paragraph text: " + paragraphs.items[0].text);
    }
});
```

---

### items

**Type:** `Word.Paragraph[]`

Gets the loaded child items in this collection.

#### Examples

**Example**: Create a new list starting from the second paragraph in the document, add list items at different positions and levels, and insert a non-list paragraph after the list.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/20-lists/insert-list.yaml

// This example starts a new list with the second paragraph.
await Word.run(async (context) => {
  const paragraphs: Word.ParagraphCollection = context.document.body.paragraphs;
  paragraphs.load("$none");

  await context.sync();

  // Start new list using the second paragraph.
  const list: Word.List = paragraphs.items[1].startNewList();
  list.load("$none");

  await context.sync();

  // To add new items to the list, use Start or End on the insertLocation parameter.
  list.insertParagraph("New list item at the start of the list", "Start");
  const paragraph: Word.Paragraph = list.insertParagraph("New list item at the end of the list (set to list level 5)", "End");

  // Set up list level for the list item.
  paragraph.listItem.level = 4;

  // To add paragraphs outside the list, use Before or After.
  list.insertParagraph("New paragraph goes after (not part of the list)", "After");

  await context.sync();
});
```

---

## Methods

### add

**Kind:** `create`

Returns a `Paragraph` object that represents a new, blank paragraph added to the document.

#### Signature

**Parameters:**
- `range`: `Word.Range` (required)
  The range before which you want the new paragraph to be added. The new paragraph doesn't replace the range.

**Returns:** `Word.Paragraph`

#### Examples

**Example**: Add a new blank paragraph at the end of the document and insert text into it

```typescript
await Word.run(async (context) => {
    // Get the document body
    const body = context.document.body;
    
    // Add a new blank paragraph at the end of the document
    const newParagraph = body.paragraphs.add();
    
    // Insert text into the new paragraph
    newParagraph.insertText("This is a new paragraph added to the document.", Word.InsertLocation.start);
    
    await context.sync();
});
```

---

### closeUp

Removes any spacing before the specified paragraphs.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Remove all spacing before paragraphs in the first content control to create a more compact layout

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document
    const contentControl = context.document.contentControls.getFirst();
    
    // Get all paragraphs within the content control
    const paragraphs = contentControl.paragraphs;
    
    // Load the paragraphs to work with them
    paragraphs.load("text");
    
    // Remove spacing before all paragraphs
    paragraphs.closeUp();
    
    await context.sync();
    
    console.log("Spacing before paragraphs removed successfully");
});
```

---

### decreaseSpacing

Decreases the spacing before and after paragraphs in six-point increments.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Decrease the spacing before and after all paragraphs in the document by six points

```typescript
await Word.run(async (context) => {
    // Get all paragraphs in the document
    const paragraphs = context.document.body.paragraphs;
    
    // Decrease spacing before and after paragraphs
    paragraphs.decreaseSpacing();
    
    await context.sync();
});
```

---

### getFirst

**Kind:** `read`

Gets the first paragraph in this collection. Throws an `ItemNotFound` error if the collection is empty.

#### Signature

**Returns:** `Word.Paragraph`

#### Examples

**Example**: Retrieve and display all annotations with their IDs, states, and critique content from the first paragraph in the current selection.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-annotations.yaml

// Gets annotations found in the selected paragraph.
await Word.run(async (context) => {
  const paragraph: Word.Paragraph = context.document.getSelection().paragraphs.getFirst();
  const annotations: Word.AnnotationCollection = paragraph.getAnnotations();
  annotations.load("id,state,critiqueAnnotation");

  await context.sync();

  console.log("Annotations found:");

  for (let i = 0; i < annotations.items.length; i++) {
    const annotation: Word.Annotation = annotations.items[i];

    console.log(`ID ${annotation.id} - state '${annotation.state}':`, annotation.critiqueAnnotation.critique);
  }
});
```

---

### getFirstOrNullObject

**Kind:** `read`

Gets the first paragraph in this collection. If the collection is empty, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

#### Signature

**Returns:** `Word.Paragraph`

#### Examples

**Example**: Check if a document has any paragraphs and display the text of the first paragraph if it exists, or show a message if the document is empty.

```typescript
await Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    const firstParagraph = paragraphs.getFirstOrNullObject();
    firstParagraph.load("text, isNullObject");
    
    await context.sync();
    
    if (firstParagraph.isNullObject) {
        console.log("The document has no paragraphs.");
    } else {
        console.log("First paragraph text: " + firstParagraph.text);
    }
});
```

---

### getLast

**Kind:** `read`

Gets the last paragraph in this collection. Throws an `ItemNotFound` error if the collection is empty.

#### Signature

**Returns:** `Word.Paragraph`

#### Examples

**Example**: Insert an inline image at the end of a new paragraph that follows the last paragraph in the document body.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/15-images/insert-and-get-pictures.yaml

// Inserts an image anchored to the last paragraph.
await Word.run(async (context) => {
  context.document.body.paragraphs
    .getLast()
    .insertParagraph("", "After")
    .insertInlinePictureFromBase64(base64Image, "End");

  await context.sync();
});
```

---

### getLastOrNullObject

**Kind:** `read`

Gets the last paragraph in this collection. If the collection is empty, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

#### Signature

**Returns:** `Word.Paragraph`

#### Examples

**Example**: Check if a document has any paragraphs and highlight the last paragraph if it exists

```typescript
await Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    const lastParagraph = paragraphs.getLastOrNullObject();
    
    lastParagraph.load("isNullObject, text");
    await context.sync();
    
    if (lastParagraph.isNullObject) {
        console.log("No paragraphs found in the document.");
    } else {
        lastParagraph.font.highlightColor = "yellow";
        console.log("Last paragraph highlighted: " + lastParagraph.text);
    }
    
    await context.sync();
});
```

---

### increaseSpacing

Increases the spacing before and after paragraphs in six-point increments.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Increase the spacing before and after all paragraphs in the document by six points

```typescript
await Word.run(async (context) => {
    // Get all paragraphs in the document
    const paragraphs = context.document.body.paragraphs;
    
    // Increase spacing before and after all paragraphs
    paragraphs.increaseSpacing();
    
    await context.sync();
});
```

---

### indent

Indents the paragraphs by one level.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Indent all paragraphs in the document by one level

```typescript
await Word.run(async (context) => {
    // Get all paragraphs in the document
    const paragraphs = context.document.body.paragraphs;
    
    // Load the paragraphs
    paragraphs.load("text");
    
    // Indent all paragraphs by one level
    paragraphs.indent();
    
    await context.sync();
    
    console.log("All paragraphs have been indented by one level");
});
```

---

### indentCharacterWidth

Indents the paragraphs in the collection by the specified number of characters.

#### Signature

**Parameters:**
- `count`: `number` (required)
  The number of characters by which the specified paragraphs are to be indented.

**Returns:** `void`

#### Examples

**Example**: Indent all paragraphs in the document body by 5 characters to the right

```typescript
await Word.run(async (context) => {
    // Get all paragraphs in the document body
    const paragraphs = context.document.body.paragraphs;
    
    // Indent all paragraphs by 5 characters
    paragraphs.indentCharacterWidth(5);
    
    await context.sync();
});
```

---

### indentFirstLineCharacterWidth

Indents the first line of the paragraphs in the collection by the specified number of characters.

#### Signature

**Parameters:**
- `count`: `number` (required)
  The number of characters by which the first line of each specified paragraph is to be indented.

**Returns:** `void`

#### Examples

**Example**: Indent the first line of all paragraphs in the document by 5 characters

```typescript
await Word.run(async (context) => {
    // Get all paragraphs in the document
    const paragraphs = context.document.body.paragraphs;
    
    // Indent the first line of each paragraph by 5 characters
    paragraphs.indentFirstLineCharacterWidth(5);
    
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
  - `options`: `Word.Interfaces.ParagraphCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.ParagraphCollection`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.ParagraphCollection`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `OfficeExtension.LoadOption` (optional)
    `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.ParagraphCollection`

#### Examples

**Example**: Retrieve all paragraphs from the document body with their text content and font size properties.

```typescript
// This example shows how to get the paragraphs in the Word document
// along with their text and font size properties.
// 
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Create a proxy object for the paragraphs collection.
    const paragraphs = context.document.body.paragraphs;

    // Queue a command to load the text and font properties.
    // It is best practice to always specify the property set. Otherwise, all properties are
    // returned in on the object.
    paragraphs.load('text, font/size');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    await context.sync();

    // Insert code that works with the paragraphs loaded by paragraphs.load().
});
```

---

### openOrCloseUp

Toggles spacing before paragraphs.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Toggle the spacing before all paragraphs in the document to add or remove extra space above them

```typescript
await Word.run(async (context) => {
    // Get all paragraphs in the document
    const paragraphs = context.document.body.paragraphs;
    
    // Load the paragraphs
    paragraphs.load("text");
    await context.sync();
    
    // Toggle spacing before all paragraphs
    paragraphs.openOrCloseUp();
    
    await context.sync();
});
```

---

### openUp

Sets spacing before the specified paragraphs to 12 points.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Add 12-point spacing before all paragraphs in the document body to improve readability

```typescript
await Word.run(async (context) => {
    // Get all paragraphs in the document body
    const paragraphs = context.document.body.paragraphs;
    
    // Set 12-point spacing before all paragraphs
    paragraphs.openUp();
    
    await context.sync();
});
```

---

### outdent

Removes one level of indent for the paragraphs.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Remove one level of indentation from all paragraphs in the document body

```typescript
await Word.run(async (context) => {
    // Get all paragraphs in the document body
    const paragraphs = context.document.body.paragraphs;
    
    // Load the paragraphs
    paragraphs.load("text");
    await context.sync();
    
    // Remove one level of indent from all paragraphs
    paragraphs.outdent();
    
    await context.sync();
});
```

---

### outlineDemote

Applies the next heading level style (Heading 1 through Heading 8) to the specified paragraphs.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Demote the first three paragraphs in the document to the next heading level (e.g., Heading 1 becomes Heading 2)

```typescript
await Word.run(async (context) => {
    // Get the first three paragraphs in the document
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("items");
    
    await context.sync();
    
    // Create a collection from the first three paragraphs
    const firstThree = paragraphs.items.slice(0, 3);
    
    // Demote each paragraph to the next heading level
    for (const paragraph of firstThree) {
        paragraph.outlineDemote();
    }
    
    await context.sync();
});
```

---

### outlineDemoteToBody

Demotes the specified paragraphs to body text by applying the Normal style.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Demote all heading paragraphs in the document to body text by applying the Normal style

```typescript
await Word.run(async (context) => {
    // Get all paragraphs in the document
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("style");
    
    await context.sync();
    
    // Filter paragraphs that are headings
    const headingParagraphs = paragraphs.items.filter(p => 
        p.style.startsWith("Heading")
    );
    
    // Create a collection from heading paragraphs
    const headingCollection = context.document.body.paragraphs;
    headingCollection.load("items");
    await context.sync();
    
    // Get only heading paragraphs
    const filteredParagraphs = headingCollection.items.filter(p => 
        headingParagraphs.includes(p)
    );
    
    // Demote all heading paragraphs to body text
    if (filteredParagraphs.length > 0) {
        const collection = context.document.body.paragraphs;
        collection.outlineDemoteToBody();
    }
    
    await context.sync();
});
```

---

### outlinePromote

Applies the previous heading level style (Heading 1 through Heading 8) to the paragraphs in the collection.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Promote all paragraphs in the document body to the previous heading level (e.g., Heading 3 becomes Heading 2)

```typescript
await Word.run(async (context) => {
    // Get all paragraphs in the document body
    const paragraphs = context.document.body.paragraphs;
    
    // Load the current style of paragraphs to verify they are headings
    paragraphs.load("style");
    await context.sync();
    
    // Promote all paragraphs to the previous heading level
    paragraphs.outlinePromote();
    
    await context.sync();
    
    console.log("All paragraphs promoted to previous heading level");
});
```

---

### space1

Sets the specified paragraphs to single spacing.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Set all paragraphs in the document to single spacing

```typescript
await Word.run(async (context) => {
    // Get all paragraphs in the document
    const paragraphs = context.document.body.paragraphs;
    
    // Set all paragraphs to single spacing
    paragraphs.space1();
    
    await context.sync();
});
```

---

### space1Pt5

Sets the specified paragraphs to 1.5-line spacing.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Set all paragraphs in the document to 1.5-line spacing

```typescript
await Word.run(async (context) => {
    // Get all paragraphs in the document
    const paragraphs = context.document.body.paragraphs;
    
    // Set 1.5-line spacing for all paragraphs
    paragraphs.space1Pt5();
    
    await context.sync();
});
```

---

### space2

Sets the specified paragraphs to double spacing.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Set all paragraphs in the document body to double spacing

```typescript
await Word.run(async (context) => {
    // Get all paragraphs in the document body
    const paragraphs = context.document.body.paragraphs;
    
    // Set all paragraphs to double spacing
    paragraphs.space2();
    
    await context.sync();
});
```

---

### tabHangingIndent

Sets a hanging indent to the specified number of tab stops.

#### Signature

**Parameters:**
- `count`: `number` (required)
  The number of tab stops to indent (if positive) or the number of tab stops to remove from the indent (if negative).

**Returns:** `void`

#### Examples

**Example**: Set a hanging indent of 2 tab stops for all paragraphs in the document body

```typescript
await Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("items");
    
    await context.sync();
    
    paragraphs.tabHangingIndent(2);
    
    await context.sync();
});
```

---

### tabIndent

Sets the left indent for the specified paragraphs to the specified number of tab stops.

#### Signature

**Parameters:**
- `count`: `number` (required)
  The number of tab stops to indent (if positive) or the number of tab stops to remove from the indent (if negative).

**Returns:** `void`

#### Examples

**Example**: Indent all paragraphs in the document by 2 tab stops from the left margin

```typescript
await Word.run(async (context) => {
    // Get all paragraphs in the document
    const paragraphs = context.document.body.paragraphs;
    
    // Indent all paragraphs by 2 tab stops
    paragraphs.tabIndent(2);
    
    await context.sync();
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.ParagraphCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.ParagraphCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

#### Signature

**Returns:** `Word.Interfaces.ParagraphCollectionData`

#### Examples

**Example**: Export paragraph data from a document to JSON format for logging or external storage

```typescript
await Word.run(async (context) => {
    // Get all paragraphs in the document
    const paragraphs = context.document.body.paragraphs;
    
    // Load properties we want to export
    paragraphs.load("text, style, alignment");
    
    await context.sync();
    
    // Convert the collection to a plain JavaScript object
    const paragraphData = paragraphs.toJSON();
    
    // Now we can use the data outside the Word context
    console.log("Paragraph data:", JSON.stringify(paragraphData, null, 2));
    
    // Access the items array
    console.log(`Total paragraphs: ${paragraphData.items.length}`);
    paragraphData.items.forEach((para, index) => {
        console.log(`Paragraph ${index + 1}: ${para.text}`);
    });
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.ParagraphCollection`

#### Examples

**Example**: Track paragraphs across multiple sync calls to maintain references when modifying their properties outside of a single sequential batch

```typescript
await Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("text");
    
    // Track the collection to use it across multiple sync calls
    paragraphs.track();
    
    await context.sync();
    
    // First sync - log paragraph count
    console.log(`Found ${paragraphs.items.length} paragraphs`);
    
    await context.sync();
    
    // Second sync - modify paragraphs (tracking prevents InvalidObjectPath error)
    for (let i = 0; i < paragraphs.items.length; i++) {
        paragraphs.items[i].font.bold = true;
    }
    
    await context.sync();
    
    // Clean up tracking when done
    paragraphs.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

#### Signature

**Returns:** `Word.ParagraphCollection`

#### Examples

**Example**: Load all paragraphs in a document, process them to get their text content, then untrack the collection to free memory after use.

```typescript
await Word.run(async (context) => {
    // Get all paragraphs in the document
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("text");
    
    await context.sync();
    
    // Process the paragraphs (e.g., log their text)
    console.log(`Found ${paragraphs.items.length} paragraphs`);
    paragraphs.items.forEach((paragraph, index) => {
        console.log(`Paragraph ${index + 1}: ${paragraph.text}`);
    });
    
    // Release memory associated with the collection
    paragraphs.untrack();
    
    await context.sync();
});
```

---

## Source

- https://docs.microsoft.com/en-us/javascript/api/word
- https://docs.microsoft.com/en-us/javascript/api/office/officeextension.clientobject
