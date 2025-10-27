# Word.Annotation

**Package:** `word`

**API Set:** WordApi 1.7

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents an annotation attached to a paragraph.

## Class Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-annotations.yaml

// Accepts the first annotation found in the selected paragraph.
await Word.run(async (context) => {
  const paragraph: Word.Paragraph = context.document.getSelection().paragraphs.getFirst();
  const annotations: Word.AnnotationCollection = paragraph.getAnnotations();
  annotations.load("id,state,critiqueAnnotation");

  await context.sync();

  for (let i = 0; i < annotations.items.length; i++) {
    const annotation: Word.Annotation = annotations.items[i];

    if (annotation.state === Word.AnnotationState.created) {
      console.log(`Accepting ID ${annotation.id}...`);
      annotation.critiqueAnnotation.accept();

      await context.sync();
      break;
    }
  }
});
```

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access an annotation's request context to verify the connection to the Office host application before performing operations on the annotation.

```typescript
await Word.run(async (context) => {
    // Get the first paragraph with annotations
    const paragraph = context.document.body.paragraphs.getFirst();
    const annotations = paragraph.getAnnotations();
    annotations.load("items");
    
    await context.sync();
    
    if (annotations.items.length > 0) {
        const annotation = annotations.items[0];
        
        // Access the annotation's request context
        const annotationContext = annotation.context;
        
        // Verify the context is valid and connected
        console.log("Annotation context is connected:", annotationContext !== null);
        
        // Use the context to perform operations
        annotation.load("id,critiqueAnnotation");
        await annotationContext.sync();
        
        console.log("Annotation ID:", annotation.id);
    }
});
```

---

### critiqueAnnotation

**Type:** `Word.CritiqueAnnotation`

**Since:** WordApi 1.7

Gets the critique annotation object.

#### Examples

**Example**: Retrieve and display all annotations from the first paragraph in the current selection, including their ID, state, and critique content.

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

### id

**Type:** `string`

**Since:** WordApi 1.7

Gets the unique identifier, which is meant to be used for easier tracking of Annotation objects.

#### Examples

**Example**: Accept the first annotation in the created state found in the selected paragraph.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-annotations.yaml

// Accepts the first annotation found in the selected paragraph.
await Word.run(async (context) => {
  const paragraph: Word.Paragraph = context.document.getSelection().paragraphs.getFirst();
  const annotations: Word.AnnotationCollection = paragraph.getAnnotations();
  annotations.load("id,state,critiqueAnnotation");

  await context.sync();

  for (let i = 0; i < annotations.items.length; i++) {
    const annotation: Word.Annotation = annotations.items[i];

    if (annotation.state === Word.AnnotationState.created) {
      console.log(`Accepting ID ${annotation.id}...`);
      annotation.critiqueAnnotation.accept();

      await context.sync();
      break;
    }
  }
});
```

---

### state

**Type:** `Word.AnnotationState | "Created" | "Accepted" | "Rejected"`

**Since:** WordApi 1.7

Gets the state of the annotation.

#### Examples

**Example**: Reject the last annotation in the selected paragraph that has a state of "created".

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-annotations.yaml

// Rejects the last annotation found in the selected paragraph.
await Word.run(async (context) => {
  const paragraph: Word.Paragraph = context.document.getSelection().paragraphs.getFirst();
  const annotations: Word.AnnotationCollection = paragraph.getAnnotations();
  annotations.load("id,state,critiqueAnnotation");

  await context.sync();

  for (let i = annotations.items.length - 1; i >= 0; i--) {
    const annotation: Word.Annotation = annotations.items[i];

    if (annotation.state === Word.AnnotationState.created) {
      console.log(`Rejecting ID ${annotation.id}...`);
      annotation.critiqueAnnotation.reject();

      await context.sync();
      break;
    }
  }
});
```

---

## Methods

### delete

**Kind:** `delete`

Deletes the annotation.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Delete all annotations from the currently selected paragraph in the document.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-annotations.yaml

// Deletes all annotations found in the selected paragraph.
await Word.run(async (context) => {
  const paragraph: Word.Paragraph = context.document.getSelection().paragraphs.getFirst();
  const annotations: Word.AnnotationCollection = paragraph.getAnnotations();
  annotations.load("id");

  await context.sync();

  const ids = [];
  for (let i = 0; i < annotations.items.length; i++) {
    const annotation: Word.Annotation = annotations.items[i];

    ids.push(annotation.id);
    annotation.delete();
  }

  await context.sync();

  console.log("Annotations deleted:", ids);
});
```

---

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.AnnotationLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.Annotation`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.Annotation`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    select is a comma-delimited string that specifies the properties to load, and expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.Annotation`

#### Examples

**Example**: Load and display the ID and state of the first annotation in the document

```typescript
await Word.run(async (context) => {
    // Get the first annotation in the document
    const annotations = context.document.body.getAnnotations();
    const firstAnnotation = annotations.getFirst();
    
    // Load specific properties of the annotation
    firstAnnotation.load("id, state");
    
    // Sync to execute the load command
    await context.sync();
    
    // Now we can read the loaded properties
    console.log("Annotation ID: " + firstAnnotation.id);
    console.log("Annotation State: " + firstAnnotation.state);
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.Annotation object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.AnnotationData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.AnnotationData`

#### Examples

**Example**: Retrieve annotation data from the first paragraph and serialize it to JSON format for logging or storage purposes.

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const paragraph = context.document.body.paragraphs.getFirst();
    
    // Get annotations from the paragraph
    const annotations = paragraph.getAnnotations();
    annotations.load("items");
    
    await context.sync();
    
    if (annotations.items.length > 0) {
        const annotation = annotations.items[0];
        
        // Load properties you want to serialize
        annotation.load("id,state,critiqueAnnotation");
        await context.sync();
        
        // Convert the annotation to a plain JavaScript object
        const annotationData = annotation.toJSON();
        
        // Now you can use JSON.stringify or log the data
        console.log(JSON.stringify(annotationData, null, 2));
        console.log("Annotation ID:", annotationData.id);
        console.log("Annotation State:", annotationData.state);
    }
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.Annotation`

#### Examples

**Example**: Track an annotation object to maintain its reference across multiple sync calls when checking and modifying its properties

```typescript
await Word.run(async (context) => {
    // Get the first paragraph with annotations
    const paragraph = context.document.body.paragraphs.getFirst();
    const annotations = paragraph.getAnnotations();
    
    await context.sync();
    
    if (annotations.items.length > 0) {
        const annotation = annotations.items[0];
        
        // Track the annotation to use it across multiple sync calls
        annotation.track();
        
        // Load properties
        annotation.load("id,state,critiqueAnnotation");
        await context.sync();
        
        // Now we can safely use the annotation in subsequent operations
        console.log("Annotation ID:", annotation.id);
        console.log("Annotation State:", annotation.state);
        
        // Perform another sync and continue using the tracked object
        await context.sync();
        
        // The tracked annotation remains valid
        if (annotation.critiqueAnnotation) {
            console.log("Critique found");
        }
        
        // Clean up tracking when done
        annotation.untrack();
    }
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.Annotation`

#### Examples

**Example**: Load an annotation, use its properties, then untrack it to free memory after you're done working with it.

```typescript
await Word.run(async (context) => {
    // Get the first annotation in the document
    const annotations = context.document.body.getAnnotations();
    const firstAnnotation = annotations.getFirstOrNullObject();
    firstAnnotation.track();
    firstAnnotation.load("id,state");
    
    await context.sync();
    
    if (!firstAnnotation.isNullObject) {
        console.log(`Annotation ID: ${firstAnnotation.id}`);
        console.log(`Annotation state: ${firstAnnotation.state}`);
        
        // Untrack the annotation to release memory
        firstAnnotation.untrack();
        await context.sync();
    }
});
```

---

## Source

- https://docs.microsoft.com/en-us/javascript/api/word
