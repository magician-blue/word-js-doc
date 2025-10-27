# Word.Source

**Package:** `word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents an individual source, such as a book, journal article, or interview.

## Properties

### context

**Type:** `None`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a Source object to verify the connection between the add-in and Word application before performing operations on bibliography sources.

```typescript
await Word.run(async (context) => {
    // Get the bibliography sources from the document
    const sources = context.document.bibliography.sources;
    sources.load("items");
    await context.sync();
    
    if (sources.items.length > 0) {
        const firstSource = sources.items[0];
        
        // Access the context property to verify the connection
        const sourceContext = firstSource.context;
        
        // Use the context to perform operations
        firstSource.load("tag, title");
        await sourceContext.sync();
        
        console.log(`Source tag: ${firstSource.tag}`);
        console.log(`Source title: ${firstSource.title}`);
        console.log("Context connection verified and operational");
    }
});
```

---

### isCited

**Type:** `None`

Gets if the Source object has been cited in the document.

#### Examples

**Example**: Check if a specific source in the bibliography has been cited anywhere in the document and display the result in the console.

```typescript
await Word.run(async (context) => {
    const sources = context.document.bibliography.sources;
    sources.load("items");
    
    await context.sync();
    
    if (sources.items.length > 0) {
        const firstSource = sources.items[0];
        firstSource.load("isCited, tag");
        
        await context.sync();
        
        console.log(`Source "${firstSource.tag}" is cited: ${firstSource.isCited}`);
    }
});
```

---

### tag

**Type:** `None`

Gets the tag of the source.

#### Examples

**Example**: Get the tag identifier of a bibliography source and display it in the document

```typescript
await Word.run(async (context) => {
    // Get the first bibliography source from the document
    const sources = context.document.bibliography.sources;
    sources.load("items");
    await context.sync();
    
    if (sources.items.length > 0) {
        const source = sources.items[0];
        
        // Get the tag of the source
        const tag = source.tag;
        
        // Insert the tag at the end of the document
        const body = context.document.body;
        body.insertParagraph(`Source tag: ${tag}`, Word.InsertLocation.end);
        
        await context.sync();
    }
});
```

---

### xml

**Type:** `None`

Gets the XML representation of the source.

#### Examples

**Example**: Retrieve and log the XML representation of a bibliography source to inspect its structure

```typescript
await Word.run(async (context) => {
    // Get the bibliography sources from the document
    const sources = context.document.bibliography.sources;
    sources.load("items");
    
    await context.sync();
    
    // Get the first source and retrieve its XML representation
    if (sources.items.length > 0) {
        const firstSource = sources.items[0];
        const sourceXml = firstSource.xml;
        
        console.log("Source XML:", sourceXml);
    } else {
        console.log("No bibliography sources found in the document");
    }
});
```

---

## Methods

### delete

**Kind:** `delete`

Deletes the Source object.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Delete the first source from the document's bibliography

```typescript
await Word.run(async (context) => {
    // Get the first source from the document
    const sources = context.document.bibliography.sources;
    sources.load("items");
    
    await context.sync();
    
    if (sources.items.length > 0) {
        const firstSource = sources.items[0];
        
        // Delete the source
        firstSource.delete();
        
        await context.sync();
        console.log("First source deleted successfully");
    } else {
        console.log("No sources found in the document");
    }
});
```

---

### getFieldByName

**Kind:** `read`

Returns the value of a field in the bibliography Source object.

#### Signature

**Parameters:**
- `name`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Retrieve and display the author and title fields from the first bibliography source in the document.

```typescript
await Word.run(async (context) => {
    const sources = context.document.bibliography.sources;
    sources.load("items");
    
    await context.sync();
    
    if (sources.items.length > 0) {
        const firstSource = sources.items[0];
        
        // Get specific fields from the source
        const author = firstSource.getFieldByName("Author");
        const title = firstSource.getFieldByName("Title");
        
        await context.sync();
        
        console.log("Author: " + author.value);
        console.log("Title: " + title.value);
    } else {
        console.log("No bibliography sources found in the document.");
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
  - `options`: `None` (required)

  **Returns:** `None`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `None` (required)

  **Returns:** `None`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `None` (required)

  **Returns:** `None`

#### Examples

**Example**: Load and read the title and tag properties of the first bibliography source in the document

```typescript
await Word.run(async (context) => {
    // Get the first source from the bibliography
    const sources = context.document.bibliography.sources;
    sources.load("items");
    await context.sync();
    
    if (sources.items.length > 0) {
        const firstSource = sources.items[0];
        
        // Load specific properties of the source
        firstSource.load("title, tag");
        await context.sync();
        
        // Now we can read the loaded properties
        console.log("Source Title: " + firstSource.title);
        console.log("Source Tag: " + firstSource.tag);
    }
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.Source object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.SourceData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.SourceData`
a plain JavaScript object (typed as Word.Interfaces.SourceData) that contains shallow copies of any loaded child properties from the original object

#### Examples

**Example**: Serialize a bibliography source to a plain JavaScript object and log it to the console for debugging or data export purposes.

```typescript
await Word.run(async (context) => {
    // Get the first source from the document's bibliography
    const sources = context.document.bibliography.sources;
    sources.load("items");
    await context.sync();
    
    if (sources.items.length > 0) {
        const firstSource = sources.items[0];
        
        // Load properties we want to include in the JSON output
        firstSource.load("tag, type, title, author, year");
        await context.sync();
        
        // Convert the source to a plain JavaScript object
        const sourceData = firstSource.toJSON();
        
        // Now we can use the plain object for logging, storage, or transmission
        console.log("Source as JSON:", JSON.stringify(sourceData, null, 2));
    }
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Track a bibliography source object across multiple sync calls to safely modify its properties without getting an "InvalidObjectPath" error

```typescript
await Word.run(async (context) => {
    const sources = context.document.bibliography.sources;
    sources.load("items");
    await context.sync();
    
    // Get the first source and track it
    const source = sources.items[0];
    source.track();
    source.load("tag, title");
    await context.sync();
    
    console.log(`Source tag: ${source.tag}`);
    
    // Can safely use the source across multiple sync calls
    // because it's being tracked
    await context.sync();
    
    console.log(`Source title: ${source.title}`);
    
    // Untrack when done to free up memory
    source.untrack();
    await context.sync();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with t

#### Signature

**Returns:** `None`

#### Examples

**Example**: Release memory for a bibliography source after reading its properties to prevent memory leaks in a long-running add-in

```typescript
await Word.run(async (context) => {
    // Get the first source from the bibliography
    const sources = context.document.bibliography.sources;
    sources.load("items");
    await context.sync();
    
    if (sources.items.length > 0) {
        const source = sources.items[0];
        source.load("title, author");
        await context.sync();
        
        // Use the source data
        console.log(`Title: ${source.title}`);
        console.log(`Author: ${source.author}`);
        
        // Release memory when done with the source
        source.untrack();
    }
});
```

---

## Source

- /en-us/javascript/api/word
- /en-us/javascript/api/office/officeextension.clientobject
- /en-us/javascript/api/requirement-sets/word/word-api-requirement-sets
