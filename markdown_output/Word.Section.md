# Word.Section

**Package:** `word`

**API Set:** WordApi 1.1 None

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents a section in a Word document.

## Class Examples

**Example**: Inserts a section break on the next page.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/insert-section-breaks.yaml

// Inserts a section break on the next page.
await Word.run(async (context) => {
  const body: Word.Body = context.document.body;
  body.insertBreak(Word.BreakType.sectionNext, Word.InsertLocation.end);

  await context.sync();

  console.log("Inserted section break on next page.");
});
```

## Properties

### body

**Type:** `Word.Body`

**Since:** WordApi 1.1

Gets the body object of the section. This doesn't include the header/footer and other section metadata.

#### Examples

**Example**: Add a paragraph with text to the body of the first section in the document

```typescript
await Word.run(async (context) => {
    // Get the first section in the document
    const firstSection = context.document.sections.getFirst();
    
    // Get the body of the section
    const sectionBody = firstSection.body;
    
    // Insert a paragraph at the end of the section body
    sectionBody.insertParagraph("This text is added to the section body.", Word.InsertLocation.end);
    
    await context.sync();
});
```

---

### borders

**Type:** `Word.BorderUniversalCollection`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns a BorderUniversalCollection object that represents all the borders in the section.

#### Examples

**Example**: Set all section borders to a solid blue line with 2.25pt width

```typescript
await Word.run(async (context) => {
    const section = context.document.sections.getFirst();
    const borders = section.borders;
    borders.load("items");
    
    await context.sync();
    
    for (let i = 0; i < borders.items.length; i++) {
        borders.items[i].type = Word.BorderType.single;
        borders.items[i].color = "blue";
        borders.items[i].width = 2.25;
    }
    
    await context.sync();
});
```

---

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the section's request context to load and read the section's body text

```typescript
await Word.run(async (context) => {
    const section = context.document.sections.getFirst();
    
    // Access the request context from the section object
    const sectionContext = section.context;
    
    // Use the context to load properties
    section.load("body/text");
    
    await sectionContext.sync();
    
    console.log("Section text:", section.body.text);
});
```

---

### pageSetup

**Type:** `Word.PageSetup`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns a PageSetup object that's associated with the section.

#### Examples

**Example**: Set the first section's page orientation to landscape and configure 1-inch margins on all sides

```typescript
await Word.run(async (context) => {
    const firstSection = context.document.sections.getFirst();
    const pageSetup = firstSection.pageSetup;
    
    pageSetup.orientation = Word.PageOrientation.landscape;
    pageSetup.topMargin = 72;    // 1 inch = 72 points
    pageSetup.bottomMargin = 72;
    pageSetup.leftMargin = 72;
    pageSetup.rightMargin = 72;
    
    await context.sync();
});
```

---

### protectedForForms

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies if the section is protected for forms.

#### Examples

**Example**: Check if the first section is protected for forms and display the result, then set the second section to be protected for forms.

```typescript
await Word.run(async (context) => {
    const sections = context.document.sections;
    sections.load("items");
    await context.sync();

    // Check if first section is protected for forms
    const firstSection = sections.items[0];
    firstSection.load("protectedForForms");
    await context.sync();
    
    console.log(`First section protected for forms: ${firstSection.protectedForForms}`);
    
    // Set second section to be protected for forms
    if (sections.items.length > 1) {
        const secondSection = sections.items[1];
        secondSection.protectedForForms = true;
        await context.sync();
        
        console.log("Second section is now protected for forms");
    }
});
```

---

## Methods

### getFooter

**Kind:** `read`

Gets one of the section's footers.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `type`: `Word.HeaderFooterType` (required)
    The type of footer to return. This value must be: 'Primary', 'FirstPage', or 'EvenPages'.

  **Returns:** `Word.Body`

**Overload 2:**

  **Parameters:**
  - `type`: `"Primary" | "FirstPage" | "EvenPages"` (required)
    The type of footer to return. This value must be: 'Primary', 'FirstPage', or 'EvenPages'.

  **Returns:** `Word.Body`

#### Examples

**Example**: Add text "This is a footer." to the primary footer of the first section and wrap it in a content control.

```typescript
// Run a batch operation against the Word object model.
await Word.run(async (context) => {
    
    // Create a proxy sectionsCollection object.
    const mySections = context.document.sections;
    
    // Queue a command to load the sections.
    mySections.load('body/style');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    await context.sync();
        
    // Create a proxy object the primary footer of the first section.
    // Note that the footer is a body object.
    const myFooter = mySections.items[0].getFooter(Word.HeaderFooterType.primary);
    
    // Queue a command to insert text at the end of the footer.
    myFooter.insertText("This is a footer.", Word.InsertLocation.end);
    
    // Queue a command to wrap the header in a content control.
    myFooter.insertContentControl();
                            
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    await context.sync();
    console.log("Added a footer to the first section.");   
});
```

**Example**: Insert a paragraph with the text "This is a primary footer." at the end of the primary footer in the first section of the document.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/25-paragraph/insert-header-and-footer.yaml

await Word.run(async (context) => {
  context.document.sections
    .getFirst()
    .getFooter("Primary")
    .insertParagraph("This is a primary footer.", "End");

  await context.sync();
});
```

---

### getHeader

**Kind:** `read`

Gets one of the section's headers.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `type`: `Word.HeaderFooterType` (required)
    The type of header to return. This value must be: 'Primary', 'FirstPage', or 'EvenPages'.

  **Returns:** `Word.Body`

**Overload 2:**

  **Parameters:**
  - `type`: `"Primary" | "FirstPage" | "EvenPages"` (required)
    The type of header to return. This value must be: 'Primary', 'FirstPage', or 'EvenPages'.

  **Returns:** `Word.Body`

#### Examples

**Example**: Insert a paragraph with the text "This is a primary header." at the end of the primary header in the first section of the document.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/25-paragraph/insert-header-and-footer.yaml

await Word.run(async (context) => {
  context.document.sections
    .getFirst()
    .getHeader(Word.HeaderFooterType.primary)
    .insertParagraph("This is a primary header.", "End");

  await context.sync();
});
```

**Example**: Add the text "This is a header." to the primary header of the first section and wrap it in a content control.

```typescript
// Run a batch operation against the Word object model.
await Word.run(async (context) => {
    
    // Create a proxy sectionsCollection object.
    const mySections = context.document.sections;
    
    // Queue a command to load the sections.
    mySections.load('body/style');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    await context.sync();
    
    // Create a proxy object the primary header of the first section.
    // Note that the header is a body object.
    const myHeader = mySections.items[0].getHeader("Primary");
    
    // Queue a command to insert text at the end of the header.
    myHeader.insertText("This is a header.", Word.InsertLocation.end);
    
    // Queue a command to wrap the header in a content control.
    myHeader.insertContentControl();
                            
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    await context.sync();
    console.log("Added a header to the first section.");
});
```

---

### getNext

**Kind:** `read`

Gets the next section. Throws an ItemNotFound error if this section is the last one.

#### Signature

**Returns:** `Word.Section`

#### Examples

**Example**: Check if the current section is followed by another section, and if so, insert text at the beginning of the next section.

```typescript
await Word.run(async (context) => {
    // Get the first section
    const firstSection = context.document.sections.getFirst();
    
    try {
        // Get the next section after the first one
        const nextSection = firstSection.getNext();
        
        // Insert text at the start of the next section
        nextSection.body.insertText("This is the second section.", Word.InsertLocation.start);
        
        await context.sync();
        console.log("Text inserted in the next section.");
    } catch (error) {
        console.log("No next section found - the first section is the last one.");
    }
});
```

---

### getNextOrNullObject

**Kind:** `read`

Gets the next section. If this section is the last one, then this method will return an object with its isNullObject property set to true. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

#### Signature

**Returns:** `Word.Section`

#### Examples

**Example**: Iterate through all sections in a document and log the index of each section until reaching the last one

```typescript
await Word.run(async (context) => {
    const firstSection = context.document.sections.getFirst();
    firstSection.load("isNullObject");
    
    let currentSection = firstSection;
    let sectionIndex = 0;
    
    while (currentSection) {
        currentSection.load("isNullObject");
        await context.sync();
        
        if (!currentSection.isNullObject) {
            console.log(`Processing section ${sectionIndex}`);
            sectionIndex++;
            
            // Get the next section
            const nextSection = currentSection.getNextOrNullObject();
            nextSection.load("isNullObject");
            await context.sync();
            
            // Check if we've reached the last section
            if (nextSection.isNullObject) {
                console.log("Reached the last section");
                break;
            }
            
            currentSection = nextSection;
        } else {
            break;
        }
    }
    
    console.log(`Total sections: ${sectionIndex}`);
});
```

---

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.SectionLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.Section`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.Section`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.Section`

#### Examples

**Example**: Load and display the body text of the first section in the document

```typescript
await Word.run(async (context) => {
    // Get the first section in the document
    const firstSection = context.document.sections.getFirst();
    
    // Load the body property of the section
    firstSection.load("body");
    
    // Sync to execute the load command
    await context.sync();
    
    // Access the loaded body and get its text
    const sectionBody = firstSection.body;
    sectionBody.load("text");
    await context.sync();
    
    console.log("Section body text:", sectionBody.text);
});
```

---

### set

**Kind:** `configure`

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `Interfaces.SectionUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.Section` (required)

  **Returns:** `void`

#### Examples

**Example**: Configure multiple properties of the first section to set different first page headers and adjust page margins

```typescript
await Word.run(async (context) => {
    const firstSection = context.document.sections.getFirst();
    
    firstSection.set({
        differentFirstPageHeaderFooter: true,
        leftMargin: 72,  // 1 inch in points
        rightMargin: 72,
        topMargin: 72,
        bottomMargin: 72
    });
    
    await context.sync();
    console.log("Section properties updated successfully");
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.Section object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.SectionData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.SectionData`

#### Examples

**Example**: Get a JSON representation of the first section's properties for logging or data export purposes

```typescript
await Word.run(async (context) => {
    // Get the first section in the document
    const firstSection = context.document.sections.getFirst();
    
    // Load properties we want to include in the JSON output
    firstSection.load("body");
    
    await context.sync();
    
    // Convert the section to a plain JavaScript object
    const sectionJSON = firstSection.toJSON();
    
    // Log or export the JSON data
    console.log("Section data:", JSON.stringify(sectionJSON, null, 2));
    
    // The JSON object can now be easily serialized, stored, or transmitted
    return sectionJSON;
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.Section`

#### Examples

**Example**: Track a section object to maintain its reference across multiple sync calls when modifying section properties

```typescript
await Word.run(async (context) => {
    // Get the first section
    const section = context.document.sections.getFirst();
    
    // Track the section to prevent InvalidObjectPath errors across sync calls
    section.track();
    
    // Load properties
    section.load("body");
    await context.sync();
    
    // Make changes to the section after first sync
    section.body.insertParagraph("This is added to the tracked section", Word.InsertLocation.start);
    await context.sync();
    
    // Continue working with the section after another sync
    section.body.font.color = "blue";
    await context.sync();
    
    // Untrack when done to free up memory
    section.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.Section`

#### Examples

**Example**: Get a section, track it for performance monitoring, perform operations, then untrack it to free memory when done

```typescript
await Word.run(async (context) => {
    // Get the first section and track it
    const section = context.document.sections.getFirst();
    section.track();
    
    // Load and use the section properties
    section.load("body");
    await context.sync();
    
    // Perform operations with the section
    section.body.insertParagraph("This is added to the tracked section.", Word.InsertLocation.start);
    await context.sync();
    
    // Untrack the section to release memory
    section.untrack();
    await context.sync();
    
    console.log("Section operations completed and memory released.");
});
```

---

## Source

- https://docs.microsoft.com/en-us/javascript/api/word/word.section
