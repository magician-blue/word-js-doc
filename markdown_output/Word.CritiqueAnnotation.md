# Word.CritiqueAnnotation

**Package:** `https://learn.microsoft.com/en-us/javascript/api/word`

**API Set:** WordApi 1.7

**Extends:** `https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientobject`

## Description

Represents an annotation wrapper around critique displayed in the document.

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

**Example**: Access the request context from a critique annotation to verify the add-in's connection to the Word host application

```typescript
await Word.run(async (context) => {
    // Get the first critique annotation in the document
    const critiques = context.document.getCritiqueAnnotations();
    critiques.load("items");
    await context.sync();

    if (critiques.items.length > 0) {
        const firstCritique = critiques.items[0];
        
        // Access the request context associated with the critique annotation
        const critiqueContext = firstCritique.context;
        
        // Use the context to perform operations
        // For example, verify it's the same context or use it for debugging
        console.log("Critique annotation context is connected:", critiqueContext !== null);
        console.log("Context matches current context:", critiqueContext === context);
    }
});
```

---

### critique

**Type:** `Word.Critique`

**Since:** WordApi 1.7

Gets the critique that was passed when the annotation was inserted.

#### Examples

**Example**: Retrieve and display all annotations from the selected paragraph, including their ID, state, and critique content.

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

### range

**Type:** `Word.Range`

**Since:** WordApi 1.7

Gets the range of text that is annotated.

#### Examples

**Example**: Get the text content from a critique annotation's range and display it in the console.

```typescript
await Word.run(async (context) => {
    // Get the first critique annotation in the document
    const annotations = context.document.getCritiqueAnnotations();
    const firstAnnotation = annotations.getFirst();
    
    // Get the range of text that is annotated
    const range = firstAnnotation.range;
    range.load("text");
    
    await context.sync();
    
    console.log("Annotated text: " + range.text);
});
```

---

## Methods

### accept

**Kind:** `write`

Accepts the critique. This will change the annotation state to `accepted`.

#### Signature

**Returns:** `void`

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

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.CritiqueAnnotationLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.CritiqueAnnotation`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.CritiqueAnnotation`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.CritiqueAnnotation`

#### Examples

**Example**: Load and display the start position of the first critique annotation in the active document.

```typescript
await Word.run(async (context) => {
    // Get the first critique annotation in the document
    const annotations = context.document.getCritiqueAnnotations();
    const firstAnnotation = annotations.getFirst();
    
    // Load the start property of the annotation
    firstAnnotation.load("start");
    
    await context.sync();
    
    // Display the start position
    console.log("Critique annotation starts at position: " + firstAnnotation.start);
});
```

---

### reject

**Kind:** `write`

Rejects the critique. This will change the annotation state to `rejected`.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Reject the last critique annotation in the selected paragraph that is in the created state.

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

### toJSON

**Kind:** `serialize`

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.CritiqueAnnotation` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.CritiqueAnnotationData`) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.CritiqueAnnotationData`

#### Examples

**Example**: Serialize a critique annotation to a plain JavaScript object and log its properties to the console for debugging purposes.

```typescript
await Word.run(async (context) => {
    // Get the first critique annotation in the document
    const critiqueAnnotations = context.document.getCritiqueAnnotations();
    critiqueAnnotations.load("items");
    await context.sync();

    if (critiqueAnnotations.items.length > 0) {
        const firstCritique = critiqueAnnotations.items[0];
        
        // Load properties before calling toJSON
        firstCritique.load("range,critiqueType");
        await context.sync();
        
        // Convert the CritiqueAnnotation to a plain JavaScript object
        const critiqueData = firstCritique.toJSON();
        
        // Log the serialized data
        console.log("Critique annotation data:", JSON.stringify(critiqueData, null, 2));
    }
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member. If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.CritiqueAnnotation`

#### Examples

**Example**: Track a critique annotation object to prevent "InvalidObjectPath" errors when accessing it across multiple sync calls while processing document critiques.

```typescript
await Word.run(async (context) => {
    // Get the first critique annotation in the document
    const critiqueAnnotations = context.document.getCritiqueAnnotations();
    context.load(critiqueAnnotations, "items");
    await context.sync();
    
    if (critiqueAnnotations.items.length > 0) {
        const firstCritique = critiqueAnnotations.items[0];
        
        // Track the object to use it across multiple sync calls
        firstCritique.track();
        
        // First sync - load properties
        context.load(firstCritique, "range");
        await context.sync();
        
        // Second sync - use the tracked object safely
        const range = firstCritique.range;
        context.load(range, "text");
        await context.sync();
        
        console.log("Critique text:", range.text);
        
        // Untrack when done to free memory
        firstCritique.untrack();
    }
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

#### Signature

**Returns:** `Word.CritiqueAnnotation`

#### Examples

**Example**: Process critique annotations in a document and release their memory after collecting their data to avoid performance degradation

```typescript
await Word.run(async (context) => {
    // Get all critique annotations in the document
    const critiqueAnnotations = context.document.getCritiqueAnnotations();
    critiqueAnnotations.load("items");
    
    await context.sync();
    
    // Process each critique annotation
    const critiques = [];
    for (let i = 0; i < critiqueAnnotations.items.length; i++) {
        const critique = critiqueAnnotations.items[i];
        critique.load("id,critiqueType");
        await context.sync();
        
        // Store the critique data
        critiques.push({
            id: critique.id,
            type: critique.critiqueType
        });
        
        // Release memory for this tracked object
        critique.untrack();
    }
    
    await context.sync();
    
    console.log(`Processed ${critiques.length} critique annotations and released their memory`);
});
```

---

## Source

- https://learn.microsoft.com/en-us/javascript/api/word
