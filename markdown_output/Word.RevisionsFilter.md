# Word.RevisionsFilter

**Package:** `word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents the current settings related to the display of reviewers' comments and revision marks in the document.

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from the RevisionsFilter object to verify the connection to the Office host application before configuring revision display settings.

```typescript
await Word.run(async (context) => {
    // Get the revisions filter for the document
    const revisionsFilter = context.document.getRevisionsFilter();
    
    // Access the request context associated with the RevisionsFilter object
    const requestContext = revisionsFilter.context;
    
    // Use the context to load properties and verify the connection
    revisionsFilter.load("showRevisions");
    await requestContext.sync();
    
    console.log("Request context is connected to Office host application");
    console.log("Show revisions setting:", revisionsFilter.showRevisions);
});
```

---

### markup

**Type:** `Word.RevisionsMarkup | "None" | "Simple" | "All"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies a RevisionsMarkup value that represents the extent of reviewer markup displayed in the document.

#### Examples

**Example**: Set the document's revision markup display to show all tracked changes and comments

```typescript
await Word.run(async (context) => {
    // Get the revisions filter for the document
    const revisionsFilter = context.document.getRevisionsFilter();
    
    // Set markup to display all revision marks and comments
    revisionsFilter.markup = Word.RevisionsMarkup.all;
    
    await context.sync();
    
    console.log("Revision markup set to display all changes");
});
```

---

### reviewers

**Type:** `Word.ReviewerCollection`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the ReviewerCollection object that represents the collection of reviewers of one or more documents.

#### Examples

**Example**: Get all reviewers who have made changes to the document and display their names in the console.

```typescript
await Word.run(async (context) => {
    // Get the revisions filter
    const revisionsFilter = context.document.getRevisionsFilter();
    
    // Get the reviewers collection
    const reviewers = revisionsFilter.reviewers;
    
    // Load the reviewer names
    reviewers.load("items/name");
    
    await context.sync();
    
    // Display all reviewer names
    console.log("Document reviewers:");
    for (let i = 0; i < reviewers.items.length; i++) {
        console.log(`- ${reviewers.items[i].name}`);
    }
});
```

---

### view

**Type:** `Word.RevisionsView | "Final" | "Original"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies a RevisionsView value that represents globally whether Word displays the original version of the document or the final version, which might have revisions and formatting changes applied.

#### Examples

**Example**: Set the document view to show the original version without revisions and formatting changes

```typescript
await Word.run(async (context) => {
    // Get the revisions filter
    const revisionsFilter = context.document.getRevisionsFilter();
    
    // Set the view to show the original document
    revisionsFilter.view = Word.RevisionsView.original;
    
    await context.sync();
    
    console.log("Document view set to original (without revisions)");
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
  - `options`: `Word.Interfaces.RevisionsFilterLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.RevisionsFilter`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.RevisionsFilter`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.RevisionsFilter`

#### Examples

**Example**: Load and display the current revision filter settings to check which types of revisions are visible in the document.

```typescript
await Word.run(async (context) => {
    // Get the revision filter from the document
    const revisionFilter = context.document.getRevisionsFilter();
    
    // Load the properties of the revision filter
    revisionFilter.load("showRevisions,showComments,showInsertions,showDeletions");
    
    // Sync to read the loaded properties
    await context.sync();
    
    // Display the current filter settings
    console.log("Show Revisions:", revisionFilter.showRevisions);
    console.log("Show Comments:", revisionFilter.showComments);
    console.log("Show Insertions:", revisionFilter.showInsertions);
    console.log("Show Deletions:", revisionFilter.showDeletions);
});
```

---

### set

**Kind:** `write`

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `Interfaces.RevisionsFilterUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.RevisionsFilter` (required)

  **Returns:** `void`

#### Examples

**Example**: Configure the revisions filter to hide all revision marks and comments from specific reviewers in the document

```typescript
await Word.run(async (context) => {
    const revisionsFilter = context.document.getRevisionsFilter();
    
    // Set multiple filter properties at once
    revisionsFilter.set({
        showRevisions: false,
        showComments: false,
        reviewers: ["John Doe", "Jane Smith"]
    });
    
    await context.sync();
    console.log("Revisions filter settings updated successfully");
});
```

---

### toggleShowAllReviewers

Shows or hides all revisions in the document that contain comments and tracked changes.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Toggle the visibility of all reviewers' revisions (comments and tracked changes) in the active document

```typescript
await Word.run(async (context) => {
    // Get the revisions filter for the document
    const revisionsFilter = context.document.getRevisionsFilter();
    
    // Toggle the display of all reviewers' revisions
    revisionsFilter.toggleShowAllReviewers();
    
    await context.sync();
    
    console.log("Toggled visibility of all reviewers' revisions");
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.RevisionsFilter object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.RevisionsFilterData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.RevisionsFilterData`
Whereas the original Word.RevisionsFilter object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.RevisionsFilterData) that contains shallow copies of any loaded child properties from the original object.

#### Examples

**Example**: Get the current revision filter settings as a plain JavaScript object and log it to the console for debugging purposes.

```typescript
await Word.run(async (context) => {
    // Get the revision filter settings
    const revisionsFilter = context.document.getRevisionsFilter();
    
    // Load the properties
    revisionsFilter.load();
    
    await context.sync();
    
    // Convert to plain JavaScript object
    const filterData = revisionsFilter.toJSON();
    
    // Log the plain object (useful for debugging or serialization)
    console.log("Revision Filter Settings:", JSON.stringify(filterData, null, 2));
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.RevisionsFilter`

#### Examples

**Example**: Track a RevisionsFilter object to maintain its reference across multiple sync calls while toggling revision display settings

```typescript
await Word.run(async (context) => {
    const revisionsFilter = context.document.getRevisionsFilter();
    
    // Track the object to use it across multiple sync calls
    revisionsFilter.track();
    
    // Load current settings
    revisionsFilter.load("showRevisions");
    await context.sync();
    
    console.log("Current show revisions: " + revisionsFilter.showRevisions);
    
    // Toggle the setting
    revisionsFilter.showRevisions = !revisionsFilter.showRevisions;
    await context.sync();
    
    console.log("Updated show revisions: " + revisionsFilter.showRevisions);
    
    // Untrack when done
    revisionsFilter.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.RevisionsFilter`

#### Examples

**Example**: Get the revisions filter settings, use them to check the current state, then untrack the object to free memory after you're done with it.

```typescript
await Word.run(async (context) => {
    // Get the revisions filter from the document
    const revisionsFilter = context.document.getRevisionsFilter();
    
    // Load properties to use the object
    revisionsFilter.load("showRevisions");
    await context.sync();
    
    // Use the revisions filter object
    console.log("Show revisions: " + revisionsFilter.showRevisions);
    
    // Untrack the object to release memory when done
    revisionsFilter.untrack();
    await context.sync();
});
```

---

## Source

- https://docs.microsoft.com/en-us/javascript/api/word/word.revisionsfilter
