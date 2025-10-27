# Word.SourceCollection

**Package:** `word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents a collection of [Word.Source](/en-us/javascript/api/word/word.source) objects.

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a SourceCollection to verify the connection to the Word host application and log context information for debugging purposes.

```typescript
await Word.run(async (context) => {
    // Get the bibliography sources collection
    const sources = context.document.body.sources;
    
    // Access the request context associated with the collection
    const requestContext = sources.context;
    
    // Use the context to load properties and sync
    sources.load("items");
    await requestContext.sync();
    
    // Log context information for debugging
    console.log("Request context is connected:", requestContext !== null);
    console.log("Number of sources in collection:", sources.items.length);
    console.log("Context debug info:", requestContext.debugInfo);
});
```

---

### items

**Type:** `Word.Source[]`

Gets the loaded child items in this collection.

#### Examples

**Example**: Retrieve all bibliography sources from the document and log their titles to the console.

```typescript
await Word.run(async (context) => {
    // Get the source collection from the document
    const sources = context.document.bibliography.sources;
    
    // Load the items property to access the array of sources
    sources.load("items");
    
    await context.sync();
    
    // Access the loaded items array and iterate through sources
    const sourceItems = sources.items;
    
    for (let i = 0; i < sourceItems.length; i++) {
        const source = sourceItems[i];
        source.load("title");
    }
    
    await context.sync();
    
    // Log each source title
    sourceItems.forEach((source, index) => {
        console.log(`Source ${index + 1}: ${source.title}`);
    });
});
```

---

## Methods

### add

**Kind:** `create`

Adds a new `Source` object to the collection.

#### Signature

**Parameters:**
- `xml`: `string` (required)
  A string containing the XML data for the source.

**Returns:** `Word.Source`
A `Source` object that was added to the collection.

#### Examples

**Example**: Add a new bibliography source for a book to the document's source collection

```typescript
await Word.run(async (context) => {
    const sources = context.document.bibliography.sources;
    
    const sourceXml = `
        <b:Source>
            <b:Tag>Smith2023</b:Tag>
            <b:SourceType>Book</b:SourceType>
            <b:Author>
                <b:Author>
                    <b:NameList>
                        <b:Person>
                            <b:Last>Smith</b:Last>
                            <b:First>John</b:First>
                        </b:Person>
                    </b:NameList>
                </b:Author>
            </b:Author>
            <b:Title>TypeScript Programming Guide</b:Title>
            <b:Year>2023</b:Year>
            <b:City>New York</b:City>
            <b:Publisher>Tech Press</b:Publisher>
        </b:Source>
    `;
    
    sources.add(sourceXml);
    
    await context.sync();
    console.log("Bibliography source added successfully");
});
```

---

### getItem

**Kind:** `read`

Gets a `Source` by its index in the collection.

#### Signature

**Parameters:**
- `index`: `number` (required)
  A number that identifies the index location of a `Source` object.

**Returns:** `Word.Source`

#### Examples

**Example**: Get the third source from the bibliography and display its title in the console.

```typescript
await Word.run(async (context) => {
    const sources = context.document.bibliography.sources;
    const thirdSource = sources.getItem(2); // Index is 0-based
    thirdSource.load("title");
    
    await context.sync();
    
    console.log("Third source title: " + thirdSource.title);
});
```

---

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.SourceCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions` (required)
    Provides options for which properties of the object to load.

  **Returns:** `Word.SourceCollection`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (required)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.SourceCollection`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `OfficeExtension.LoadOption` (required)
    `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.SourceCollection`

#### Examples

**Example**: Load and display the titles of all bibliography sources in the document

```typescript
await Word.run(async (context) => {
    // Get the collection of bibliography sources
    const sources = context.document.bibliography.sources;
    
    // Load the 'title' property for all sources in the collection
    sources.load("title");
    
    // Synchronize to execute the load command
    await context.sync();
    
    // Display the titles
    console.log(`Found ${sources.items.length} sources:`);
    sources.items.forEach((source, index) => {
        console.log(`${index + 1}. ${source.title}`);
    });
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify()`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.SourceCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.SourceCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

#### Signature

**Returns:** `Word.Interfaces.SourceCollectionData`

#### Examples

**Example**: Serialize a collection of bibliography sources to JSON format for logging or external storage

```typescript
await Word.run(async (context) => {
    // Get the bibliography sources collection
    const sources = context.document.bibliography.sources;
    
    // Load properties needed for serialization
    sources.load("tag,title,author");
    
    await context.sync();
    
    // Convert the collection to a plain JavaScript object
    const sourcesJSON = sources.toJSON();
    
    // Log the serialized data (contains an "items" array with source properties)
    console.log(JSON.stringify(sourcesJSON, null, 2));
    
    // Example: Save to external storage or send to a server
    // await fetch('/api/save-sources', {
    //     method: 'POST',
    //     body: JSON.stringify(sourcesJSON)
    // });
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.SourceCollection`

#### Examples

**Example**: Track a bibliography source collection to maintain references across multiple sync calls when updating citation properties

```typescript
await Word.run(async (context) => {
    // Get the bibliography sources
    const sources = context.document.bibliography.sources;
    sources.load("items");
    await context.sync();
    
    // Track the collection to use it across multiple sync calls
    sources.track();
    
    // First sync - load source properties
    for (let i = 0; i < sources.items.length; i++) {
        sources.items[i].load("tag,title");
    }
    await context.sync();
    
    // Second sync - modify sources (tracking prevents InvalidObjectPath errors)
    console.log("Bibliography sources:");
    for (let i = 0; i < sources.items.length; i++) {
        console.log(`Tag: ${sources.items[i].tag}, Title: ${sources.items[i].title}`);
    }
    await context.sync();
    
    // Untrack when done
    sources.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

#### Signature

**Returns:** `Word.SourceCollection`

#### Examples

**Example**: Load bibliography sources, process them, then untrack the collection to free memory after use

```typescript
await Word.run(async (context) => {
    // Get the bibliography sources collection
    const sources = context.document.bibliography.sources;
    
    // Load the sources and track the collection
    sources.load("tag");
    
    await context.sync();
    
    // Process the sources (e.g., log their tags)
    console.log(`Found ${sources.items.length} sources`);
    sources.items.forEach(source => {
        console.log(source.tag);
    });
    
    // Untrack the collection to release memory
    sources.untrack();
    
    await context.sync();
});
```

---

## Source

- /en-us/javascript/api/word
