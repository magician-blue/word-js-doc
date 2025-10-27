# Word.AnnotationCollection

**Package:** `word`

**API Set:** WordApi 1.7

**Extends:** `OfficeExtension.ClientObject`

## Description

Contains a collection of [Word.Annotation](/en-us/javascript/api/word/word.annotation) objects.

## Class Examples

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

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from an AnnotationCollection to verify the connection between the add-in and Word before performing operations on annotations.

```typescript
await Word.run(async (context) => {
    const annotationCollection = context.document.body.getAnnotations();
    
    // Access the request context associated with the annotation collection
    const requestContext = annotationCollection.context;
    
    // Verify the context is valid by checking if it matches the current context
    if (requestContext === context) {
        console.log("AnnotationCollection is connected to the current Word context");
        
        // Load and sync using the context
        annotationCollection.load("items");
        await context.sync();
        
        console.log(`Found ${annotationCollection.items.length} annotations`);
    }
});
```

---

### items

**Type:** `Word.Annotation[]`

Gets the loaded child items in this collection.

#### Examples

**Example**: Get all annotations in the document and log their IDs to the console

```typescript
await Word.run(async (context) => {
    const annotations = context.document.body.getAnnotations();
    annotations.load("items");
    
    await context.sync();
    
    const annotationItems = annotations.items;
    console.log(`Found ${annotationItems.length} annotations`);
    
    annotationItems.forEach((annotation, index) => {
        console.log(`Annotation ${index + 1} ID: ${annotation.id}`);
    });
});
```

---

## Methods

### getFirst

**Kind:** `read`

Gets the first annotation in this collection. Throws an ItemNotFound error if this collection is empty.

#### Signature

**Returns:** `Word.Annotation`

#### Examples

**Example**: Get the first annotation in the document and display its ID in the console

```typescript
await Word.run(async (context) => {
    // Get the annotations collection from the document body
    const annotations = context.document.body.getAnnotations();
    
    // Get the first annotation in the collection
    const firstAnnotation = annotations.getFirst();
    
    // Load the ID property
    firstAnnotation.load("id");
    
    // Sync to execute the queued commands
    await context.sync();
    
    // Display the first annotation's ID
    console.log("First annotation ID:", firstAnnotation.id);
});
```

---

### getFirstOrNullObject

**Kind:** `read`

Gets the first annotation in this collection. If this collection is empty, then this method will return an object with its isNullObject property set to true. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

#### Signature

**Returns:** `Word.Annotation`

#### Examples

**Example**: Check if a document has any annotations and display the first annotation's ID, or show a message if no annotations exist.

```typescript
await Word.run(async (context) => {
    const annotations = context.document.body.getAnnotations();
    const firstAnnotation = annotations.getFirstOrNullObject();
    
    firstAnnotation.load("id, isNullObject");
    await context.sync();
    
    if (firstAnnotation.isNullObject) {
        console.log("No annotations found in the document.");
    } else {
        console.log(`First annotation ID: ${firstAnnotation.id}`);
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
  - `options`: `Word.Interfaces.AnnotationCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.AnnotationCollection`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.AnnotationCollection`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `OfficeExtension.LoadOption` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.AnnotationCollection`

#### Examples

**Example**: Load and display the text content of all annotations in the document

```typescript
await Word.run(async (context) => {
    // Get the annotation collection from the document body
    const annotations = context.document.body.getAnnotations();
    
    // Load the critiqueAnnotation property which contains the text
    annotations.load("items/critiqueAnnotation");
    
    await context.sync();
    
    // Display the annotation text
    annotations.items.forEach((annotation, index) => {
        console.log(`Annotation ${index + 1}: ${annotation.critiqueAnnotation}`);
    });
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.AnnotationCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.AnnotationCollectionData) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

#### Signature

**Returns:** `Word.Interfaces.AnnotationCollectionData`

#### Examples

**Example**: Retrieve all annotations from the document and export them as a plain JSON object for logging or external storage.

```typescript
await Word.run(async (context) => {
    // Get all annotations in the document
    const annotations = context.document.getAnnotations();
    
    // Load properties needed for the annotations
    annotations.load("id, state, critiqueAnnotation");
    
    await context.sync();
    
    // Convert the AnnotationCollection to a plain JavaScript object
    const annotationsData = annotations.toJSON();
    
    // The result contains an "items" array with annotation data
    console.log("Annotations as JSON:", JSON.stringify(annotationsData, null, 2));
    console.log("Number of annotations:", annotationsData.items.length);
    
    // You can now work with the plain JavaScript object
    annotationsData.items.forEach((annotation, index) => {
        console.log(`Annotation ${index + 1}: ID=${annotation.id}, State=${annotation.state}`);
    });
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.AnnotationCollection`

#### Examples

**Example**: Track an annotation collection across multiple sync calls to monitor and work with annotations without getting "InvalidObjectPath" errors

```typescript
await Word.run(async (context) => {
    const annotations = context.document.body.getAnnotations();
    
    // Track the collection to use it across multiple sync calls
    annotations.track();
    
    // First sync to load the collection
    await context.sync();
    
    console.log(`Found ${annotations.items.length} annotations`);
    
    // Can safely use the tracked collection in subsequent operations
    annotations.load("items");
    await context.sync();
    
    // Process annotations across multiple sync calls
    for (const annotation of annotations.items) {
        annotation.load("critiqueAnnotation");
        await context.sync();
        console.log(`Annotation ID: ${annotation.id}`);
    }
    
    // Untrack when done to free up memory
    annotations.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.AnnotationCollection`

#### Examples

**Example**: Load annotations from a document, process them, then untrack the collection to free memory after use.

```typescript
await Word.run(async (context) => {
    // Get the annotations collection from the document
    const annotations = context.document.body.getAnnotations();
    
    // Track the collection for change tracking
    annotations.track();
    
    // Load properties to work with
    annotations.load("items");
    await context.sync();
    
    // Process the annotations (e.g., log count)
    console.log(`Found ${annotations.items.length} annotations`);
    
    // Untrack the collection to release memory
    annotations.untrack();
    await context.sync();
});
```

---

## Source

- https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-annotations.yaml
