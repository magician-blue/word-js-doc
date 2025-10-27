# Word.Reviewer

**Package:** `word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents a single reviewer of a document in which changes have been tracked. The Reviewer object is a member of the Word.ReviewerCollection object.

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a Reviewer object to verify the connection between the add-in and Word application, then use it to load and display the reviewer's name.

```typescript
await Word.run(async (context) => {
    // Get the first reviewer from the document
    const reviewers = context.document.getReviewers();
    const firstReviewer = reviewers.getFirst();
    
    // Access the reviewer's context property
    const reviewerContext = firstReviewer.context;
    
    // Use the context to load the reviewer's properties
    firstReviewer.load("name");
    
    await reviewerContext.sync();
    
    // Display the reviewer's name
    console.log(`Reviewer name: ${firstReviewer.name}`);
    console.log(`Context is connected: ${reviewerContext !== null}`);
});
```

---

### isVisible

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies if the Reviewer object is visible.

#### Examples

**Example**: Hide all changes made by the first reviewer in the document's tracked changes

```typescript
await Word.run(async (context) => {
    const reviewers = context.document.getReviewers();
    reviewers.load("items");
    await context.sync();
    
    if (reviewers.items.length > 0) {
        const firstReviewer = reviewers.items[0];
        firstReviewer.isVisible = false;
        await context.sync();
        
        console.log("First reviewer's changes are now hidden");
    }
});
```

---

## Methods

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.ReviewerLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.Reviewer`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.Reviewer`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string }` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.Reviewer`

#### Examples

**Example**: Load and display the name of the first reviewer who has made tracked changes in the document.

```typescript
await Word.run(async (context) => {
    // Get the collection of reviewers
    const reviewers = context.document.getReviewers();
    
    // Get the first reviewer
    const firstReviewer = reviewers.getFirst();
    
    // Load the name property of the reviewer
    firstReviewer.load("name");
    
    // Sync to execute the load command
    await context.sync();
    
    // Display the reviewer's name
    console.log(`First reviewer: ${firstReviewer.name}`);
});
```

---

### set

**Kind:** `write`

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `Interfaces.ReviewerUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.Reviewer` (required)

  **Returns:** `void`

#### Examples

**Example**: Update multiple properties of a reviewer object at once, setting both the name and email address of the first reviewer in the document's tracked changes.

```typescript
await Word.run(async (context) => {
    // Get the first reviewer from the document
    const reviewers = context.document.getReviewers();
    const firstReviewer = reviewers.getFirst();
    
    // Set multiple properties at once
    firstReviewer.set({
        name: "Jane Smith",
        email: "jane.smith@example.com"
    });
    
    await context.sync();
    
    console.log("Reviewer properties updated successfully");
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.Reviewer object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.ReviewerData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.ReviewerData`

#### Examples

**Example**: Serialize a reviewer object to JSON format to log or store reviewer information from a tracked changes document.

```typescript
await Word.run(async (context) => {
    // Get the first reviewer from the document
    const reviewers = context.document.getReviewers();
    const firstReviewer = reviewers.getFirst();
    
    // Load properties needed for serialization
    firstReviewer.load("name, email, isActive");
    
    await context.sync();
    
    // Convert the reviewer object to a plain JavaScript object
    const reviewerData = firstReviewer.toJSON();
    
    // Now you can use the plain object (e.g., log it, store it, send it to a server)
    console.log("Reviewer as JSON:", JSON.stringify(reviewerData, null, 2));
    console.log("Reviewer name:", reviewerData.name);
    console.log("Reviewer email:", reviewerData.email);
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.Reviewer`

#### Examples

**Example**: Track a reviewer object across multiple sync calls to safely access its properties without getting an "InvalidObjectPath" error when working with tracked changes in a document.

```typescript
await Word.run(async (context) => {
    // Get the first reviewer from the document
    const reviewers = context.document.getReviewers();
    const firstReviewer = reviewers.getFirst();
    
    // Track the reviewer object for use across sync calls
    firstReviewer.track();
    
    // Load properties
    firstReviewer.load("name");
    
    // First sync
    await context.sync();
    
    console.log("Reviewer name: " + firstReviewer.name);
    
    // Can safely use the reviewer object after another sync
    // because it's being tracked
    await context.sync();
    
    console.log("Still accessible: " + firstReviewer.name);
    
    // Untrack when done to release memory
    firstReviewer.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.Reviewer`

#### Examples

**Example**: Track a reviewer object to monitor it, then untrack it to release memory after retrieving its properties

```typescript
await Word.run(async (context) => {
    // Get the first reviewer from the document
    const reviewers = context.document.getReviewers();
    const firstReviewer = reviewers.getFirst();
    
    // Track the reviewer object for monitoring
    firstReviewer.track();
    
    // Load and use the reviewer's properties
    firstReviewer.load("name");
    await context.sync();
    
    console.log("Reviewer name: " + firstReviewer.name);
    
    // Untrack the reviewer to release memory
    firstReviewer.untrack();
    await context.sync();
    
    console.log("Reviewer object untracked and memory released");
});
```

---

## Source

- /en-us/javascript/api/word/word.reviewer
