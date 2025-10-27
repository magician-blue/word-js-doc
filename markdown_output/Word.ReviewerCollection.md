# Word.ReviewerCollection

**Package:** `word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `officeextension.clientobject`

## Description

A collection of Word.Reviewer objects that represents the reviewers of one or more documents. The ReviewerCollection object contains the names of all reviewers who have reviewed documents opened or edited on a computer.

## Properties

### context

**Type:** `None`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the reviewer collection's context to verify the connection to the Office host application and log its debug information.

```typescript
await Word.run(async (context) => {
    const reviewerCollection = context.document.getReviewerCollection();
    reviewerCollection.load("items");
    await context.sync();
    
    // Access the context property to get request context information
    const requestContext = reviewerCollection.context;
    
    // Use the context to check if it's properly connected
    console.log("Context debug info:", requestContext.debugInfo);
    console.log("Number of reviewers:", reviewerCollection.items.length);
});
```

---

### items

**Type:** `None`

Gets the loaded child items in this collection.

#### Examples

**Example**: Get and display all reviewer names from the current document's reviewer collection.

```typescript
await Word.run(async (context) => {
    // Get the reviewer collection
    const reviewerCollection = context.document.getReviewerCollection();
    
    // Load the items property to access individual reviewers
    reviewerCollection.load("items");
    
    await context.sync();
    
    // Access the loaded items and display reviewer names
    const reviewers = reviewerCollection.items;
    console.log(`Total reviewers: ${reviewers.length}`);
    
    reviewers.forEach((reviewer, index) => {
        reviewer.load("name");
    });
    
    await context.sync();
    
    reviewers.forEach((reviewer, index) => {
        console.log(`Reviewer ${index + 1}: ${reviewer.name}`);
    });
});
```

---

## Methods

### getItem

**Kind:** `read`

Returns a Reviewer object that represents the specified item in the collection.

#### Signature

**Parameters:**
- `index`: `None` (required)

**Returns:** `Reviewer`

#### Examples

**Example**: Get the first reviewer from the collection and display their name in the console.

```typescript
await Word.run(async (context) => {
    const reviewers = context.document.getReviewerCollection();
    const firstReviewer = reviewers.getItem(0);
    firstReviewer.load("name");
    
    await context.sync();
    
    console.log("First reviewer: " + firstReviewer.name);
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

**Example**: Load and display the names of all reviewers who have made changes to the document

```typescript
await Word.run(async (context) => {
    // Get the reviewer collection
    const reviewers = context.document.getReviewerCollection();
    
    // Load the reviewer names
    reviewers.load("items/name");
    
    await context.sync();
    
    // Display the reviewer names
    console.log("Document reviewers:");
    reviewers.items.forEach(reviewer => {
        console.log(`- ${reviewer.name}`);
    });
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.ReviewerCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.ReviewerCollectionData) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

#### Signature

**Returns:** `Word.Interfaces.ReviewerCollectionData`

#### Examples

**Example**: Export reviewer information to JSON format for logging or external storage purposes.

```typescript
await Word.run(async (context) => {
    // Get the collection of reviewers
    const reviewers = context.document.getReviewers();
    
    // Load the properties we want to export
    reviewers.load("items");
    
    await context.sync();
    
    // Convert the ReviewerCollection to a plain JavaScript object
    const reviewersJSON = reviewers.toJSON();
    
    // Log the JSON output (could also send to API, save to file, etc.)
    console.log("Reviewers data:", JSON.stringify(reviewersJSON, null, 2));
    
    // Access the items array from the JSON object
    console.log(`Total reviewers: ${reviewersJSON.items.length}`);
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Track all reviewers in the document to maintain references across multiple sync calls while processing reviewer information

```typescript
await Word.run(async (context) => {
    const reviewers = context.document.getReviewerCollection();
    reviewers.load("items");
    await context.sync();
    
    // Track the collection to use it across multiple sync calls
    reviewers.track();
    
    // Perform operations across multiple syncs
    console.log(`Found ${reviewers.items.length} reviewers`);
    await context.sync();
    
    // The tracked object remains valid for further operations
    for (const reviewer of reviewers.items) {
        reviewer.load("name");
    }
    await context.sync();
    
    reviewers.items.forEach(reviewer => {
        console.log(`Reviewer: ${reviewer.name}`);
    });
    
    // Untrack when done to free up memory
    reviewers.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this

#### Signature

**Returns:** `None`

#### Examples

**Example**: Release memory for a ReviewerCollection object after retrieving and displaying reviewer information to prevent memory leaks

```typescript
await Word.run(async (context) => {
    // Get the reviewer collection
    const reviewers = context.document.getReviewerCollection();
    
    // Load the reviewer properties
    reviewers.load("items");
    await context.sync();
    
    // Use the reviewer data (e.g., log reviewer names)
    console.log(`Found ${reviewers.items.length} reviewers`);
    reviewers.items.forEach(reviewer => {
        console.log(reviewer.name);
    });
    
    // Release memory associated with the ReviewerCollection
    reviewers.untrack();
});
```

---

## Source

- https://learn.microsoft.com/en-us/javascript/api/word/word.reviewercollection
