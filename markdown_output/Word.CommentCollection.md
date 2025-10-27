# Word.CommentCollection

**Package:** `word`

**API Set:** WordApi 1.4

**Extends:** `OfficeExtension.ClientObject`

## Description

Contains a collection of [Word.Comment](/en-us/javascript/api/word/word.comment) objects.

## Class Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-comments.yaml

// Replies to the first active comment in the selected content.
await Word.run(async (context) => {
  const text = (document.getElementById("reply-text") as HTMLInputElement).value;
  const comments: Word.CommentCollection = context.document.getSelection().getComments();
  comments.load("items");
  await context.sync();

  const firstActiveComment: Word.Comment = comments.items.find((item) => item.resolved !== true);
  if (firstActiveComment) {
    const reply: Word.CommentReply = firstActiveComment.reply(text);
    console.log("Reply added.");
  } else {
    console.warn("No active comment was found in the selection, so couldn't reply.");
  }
});
```

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a CommentCollection to verify the connection between the add-in and Word application before performing operations on comments.

```typescript
await Word.run(async (context) => {
    const comments = context.document.body.getComments();
    
    // Access the request context associated with the CommentCollection
    const requestContext = comments.context;
    
    // Verify the context is valid by using it to load properties
    comments.load("items");
    await requestContext.sync();
    
    console.log(`Successfully accessed context. Found ${comments.items.length} comments.`);
});
```

---

### items

**Type:** `Word.Comment[]`

Gets the loaded child items in this collection.

#### Examples

**Example**: Reply to the first unresolved comment in the selected content with text from an input field.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-comments.yaml

// Replies to the first active comment in the selected content.
await Word.run(async (context) => {
  const text = (document.getElementById("reply-text") as HTMLInputElement).value;
  const comments: Word.CommentCollection = context.document.getSelection().getComments();
  comments.load("items");
  await context.sync();

  const firstActiveComment: Word.Comment = comments.items.find((item) => item.resolved !== true);
  if (firstActiveComment) {
    const reply: Word.CommentReply = firstActiveComment.reply(text);
    console.log("Reply added.");
  } else {
    console.warn("No active comment was found in the selection, so couldn't reply.");
  }
});
```

---

## Methods

### getFirst

**Kind:** `read`

Gets the first comment in the collection. Throws an ItemNotFound error if this collection is empty.

#### Signature

**Returns:** `Word.Comment`

#### Examples

**Example**: Get the first comment in the document and display its content in the console.

```typescript
await Word.run(async (context) => {
    const comments = context.document.body.getComments();
    const firstComment = comments.getFirst();
    firstComment.load("content");
    
    await context.sync();
    
    console.log("First comment content: " + firstComment.content);
});
```

---

### getFirstOrNullObject

**Kind:** `read`

Gets the first comment in the collection. If the collection is empty, returns an object with its isNullObject property set to true. For further information, see [*OrNullObject methods and properties*](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

#### Signature

**Returns:** `Word.Comment`

#### Examples

**Example**: Retrieve and display the text range location and content range of the first comment in the selected content.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-comments.yaml

// Gets the range of the first comment in the selected content.
await Word.run(async (context) => {
  const comment: Word.Comment = context.document.getSelection().getComments().getFirstOrNullObject();
  comment.load("contentRange");
  const range: Word.Range = comment.getRange();
  range.load("text");
  await context.sync();

  if (comment.isNullObject) {
    console.warn("No comments in the selection, so no range to get.");
    return;
  }

  console.log(`Comment location: ${range.text}`);
  const contentRange: Word.CommentContentRange = comment.contentRange;
  console.log("Comment content range:", contentRange);
});
```

---

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.CommentCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.CommentCollection`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.CommentCollection`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `OfficeExtension.LoadOption` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.CommentCollection`

#### Examples

**Example**: Load and display the author names of all comments in the document

```typescript
await Word.run(async (context) => {
    // Get the comment collection from the document
    const comments = context.document.body.getComments();
    
    // Load the author property for all comments in the collection
    comments.load("author");
    
    // Synchronize to execute the load command
    await context.sync();
    
    // Display the author of each comment
    console.log(`Found ${comments.items.length} comments`);
    comments.items.forEach((comment, index) => {
        console.log(`Comment ${index + 1} author: ${comment.author}`);
    });
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method to provide more useful output when an API object is passed to JSON.stringify(). Returns a plain JavaScript object (typed as Word.Interfaces.CommentCollectionData) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

#### Signature

**Returns:** `Word.Interfaces.CommentCollectionData`

#### Examples

**Example**: Get all comments in the document and export them as a JSON string to log their content and author information.

```typescript
await Word.run(async (context) => {
    // Get all comments in the document
    const comments = context.document.body.getComments();
    
    // Load properties we want to include in the JSON output
    comments.load("content, authorName, creationDate");
    
    await context.sync();
    
    // Convert the comment collection to a plain JavaScript object
    const commentsJSON = comments.toJSON();
    
    // Convert to JSON string and log it
    console.log(JSON.stringify(commentsJSON, null, 2));
    
    // The output will contain an "items" array with all loaded comment properties
    console.log(`Total comments: ${commentsJSON.items.length}`);
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. Shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If using this object across .sync calls and outside the sequential execution of a ".run" batch and you get an "InvalidObjectPath" error, add the object to the tracked object collection when first created. If this object is part of a collection, also track the parent collection.

#### Signature

**Returns:** `Word.CommentCollection`

#### Examples

**Example**: Track a comment collection to maintain references across multiple sync calls when monitoring and displaying comment counts

```typescript
await Word.run(async (context) => {
    const comments = context.document.body.getComments();
    comments.load("items");
    
    // Track the collection to use it across multiple sync calls
    comments.track();
    
    await context.sync();
    
    console.log(`Initial comment count: ${comments.items.length}`);
    
    // Perform additional operations that might modify the document
    // The tracked collection remains valid across sync calls
    await context.sync();
    
    console.log(`Comment count after sync: ${comments.items.length}`);
    
    // Untrack when done to free up memory
    comments.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object if previously tracked. Shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so free any objects you add once you're done using them. Call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.CommentCollection`

#### Examples

**Example**: Get all comments in a document, process them, then untrack the collection to free memory after use

```typescript
await Word.run(async (context) => {
    // Load the comments collection
    const comments = context.document.body.getComments();
    comments.load("items");
    
    await context.sync();
    
    // Process the comments (e.g., log their count)
    console.log(`Total comments: ${comments.items.length}`);
    
    // Untrack the collection to release memory
    comments.untrack();
    
    await context.sync();
});
```

---

## Source

- https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-comments.yaml
