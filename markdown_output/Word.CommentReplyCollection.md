# Word.CommentReplyCollection

**Package:** `word`

**API Set:** WordApi 1.4 None

**Extends:** `OfficeExtension.ClientObject`

## Description

Contains a collection of https://learn.microsoft.com/en-us/javascript/api/word/word.commentreply objects. Represents all comment replies in one comment thread.

## Class Examples

**Example**: Gets the replies to the first comment in the selected content.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-comments.yaml

// Gets the replies to the first comment in the selected content.
await Word.run(async (context) => {
  const comment: Word.Comment = context.document.getSelection().getComments().getFirstOrNullObject();
  comment.load("replies");
  await context.sync();

  if (comment.isNullObject) {
    console.warn("No comments in the selection, so no replies to get.");
    return;
  }

  const replies: Word.CommentReplyCollection = comment.replies;
  console.log("Replies to the first comment:", replies);
});
```

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a CommentReplyCollection to verify the connection between the add-in and Word before performing operations on comment replies.

```typescript
await Word.run(async (context) => {
    // Get the first comment in the document
    const comments = context.document.getComments();
    comments.load("items");
    await context.sync();
    
    if (comments.items.length > 0) {
        const firstComment = comments.items[0];
        const replies = firstComment.getCommentReplies();
        
        // Access the request context from the CommentReplyCollection
        const replyContext = replies.context;
        
        // Verify the context is valid and connected to the same Word context
        if (replyContext === context) {
            console.log("CommentReplyCollection is properly connected to the Word context");
            
            // Now safe to perform operations on the collection
            replies.load("items");
            await context.sync();
            console.log(`Found ${replies.items.length} replies in the comment thread`);
        }
    }
});
```

---

### items

**Type:** `Word.CommentReply[]`

Gets the loaded child items in this collection.

#### Examples

**Example**: Get all comment replies from the first comment in the document and log their content to the console.

```typescript
await Word.run(async (context) => {
    // Get the first comment in the document
    const comments = context.document.getCommentByIdOrNullObject("comment1");
    const commentReplies = comments.replies;
    
    // Load the items property to access the array of comment replies
    commentReplies.load("items");
    
    await context.sync();
    
    // Access the loaded comment reply items
    const replyItems = commentReplies.items;
    
    // Log each reply's content
    for (let i = 0; i < replyItems.length; i++) {
        replyItems[i].load("content");
    }
    
    await context.sync();
    
    replyItems.forEach((reply, index) => {
        console.log(`Reply ${index + 1}: ${reply.content}`);
    });
});
```

---

## Methods

### getFirst

**Kind:** `read`

Gets the first comment reply in the collection. Throws an ItemNotFound error if this collection is empty.

#### Signature

**Returns:** `Word.CommentReply`

#### Examples

**Example**: Get and highlight the first reply in a comment thread to mark it as the initial response

```typescript
await Word.run(async (context) => {
    // Get the first comment in the document
    const comments = context.document.body.getComments();
    const firstComment = comments.getFirst();
    
    // Get all replies for this comment
    const replies = firstComment.replies;
    
    // Get the first reply in the collection
    const firstReply = replies.getFirst();
    firstReply.load("content, authorName");
    
    await context.sync();
    
    // Log the first reply details
    console.log(`First reply by ${firstReply.authorName}: ${firstReply.content}`);
});
```

---

### getFirstOrNullObject

**Kind:** `read`

Gets the first comment reply in the collection. If the collection is empty, then this method will return an object with its isNullObject property set to true. For further information, see https://learn.microsoft.com/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties *OrNullObject methods and properties.

#### Signature

**Returns:** `Word.CommentReply`

#### Examples

**Example**: Check if a comment thread has any replies and display the first reply's content, or show a message if there are no replies.

```typescript
await Word.run(async (context) => {
    // Get the first comment in the document
    const comments = context.document.body.getComments();
    const firstComment = comments.getFirst();
    
    // Get the reply collection for this comment
    const replies = firstComment.replies;
    
    // Get the first reply or null if none exist
    const firstReply = replies.getFirstOrNullObject();
    
    // Load properties
    firstReply.load("isNullObject, content");
    
    await context.sync();
    
    // Check if a reply exists
    if (firstReply.isNullObject) {
        console.log("This comment has no replies yet.");
    } else {
        console.log("First reply content: " + firstReply.content);
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
  - `options`: `Word.Interfaces.CommentReplyCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.CommentReplyCollection`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.CommentReplyCollection`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `OfficeExtension.LoadOption` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.CommentReplyCollection`

#### Examples

**Example**: Load and display the text content of all replies in the first comment thread of the document

```typescript
await Word.run(async (context) => {
    // Get the first comment in the document
    const comments = context.document.body.getComments();
    const firstComment = comments.getFirst();
    
    // Get the reply collection for this comment
    const replies = firstComment.replies;
    
    // Load the content property for all replies in the collection
    replies.load("items/content");
    
    await context.sync();
    
    // Display the reply contents
    console.log(`Found ${replies.items.length} replies:`);
    replies.items.forEach((reply, index) => {
        console.log(`Reply ${index + 1}: ${reply.content}`);
    });
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.CommentReplyCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.CommentReplyCollectionData) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

#### Signature

**Returns:** `Word.Interfaces.CommentReplyCollectionData`

#### Examples

**Example**: Export all comment replies from the first comment in the document to a JSON string for logging or external storage

```typescript
await Word.run(async (context) => {
    // Get the first comment in the document
    const comments = context.document.getCommentByIdOrNullObject("comment1");
    const replies = comments.replies;
    
    // Load the reply collection properties
    replies.load("items");
    
    await context.sync();
    
    // Convert the comment reply collection to a plain JavaScript object
    const repliesData = replies.toJSON();
    
    // Convert to JSON string for logging or storage
    const jsonString = JSON.stringify(repliesData, null, 2);
    console.log("Comment Replies JSON:", jsonString);
    
    // The repliesData object contains an "items" array with all loaded properties
    console.log(`Number of replies: ${repliesData.items.length}`);
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.CommentReplyCollection`

#### Examples

**Example**: Track comment replies in a collection to maintain references across multiple sync calls when processing replies from different comment threads

```typescript
await Word.run(async (context) => {
    // Get the first comment in the document
    const comments = context.document.body.getComments();
    comments.load("items");
    await context.sync();
    
    if (comments.items.length > 0) {
        const firstComment = comments.items[0];
        
        // Get the comment replies collection
        const replies = firstComment.replies;
        replies.load("items");
        
        // Track the replies collection to maintain reference across sync calls
        replies.track();
        
        await context.sync();
        
        // Now we can safely work with the replies across multiple sync operations
        console.log(`Found ${replies.items.length} replies`);
        
        // Perform additional operations that require another sync
        for (const reply of replies.items) {
            reply.load("content");
        }
        
        await context.sync();
        
        // The tracked collection remains valid
        replies.items.forEach(reply => {
            console.log(`Reply content: ${reply.content}`);
        });
        
        // Untrack when done to free up memory
        replies.untrack();
    }
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.CommentReplyCollection`

#### Examples

**Example**: Load comment replies from a comment thread, process them, then untrack the collection to free memory after use.

```typescript
await Word.run(async (context) => {
    // Get the first comment in the document
    const comments = context.document.body.getComments();
    comments.load("items");
    await context.sync();
    
    if (comments.items.length > 0) {
        const firstComment = comments.items[0];
        
        // Get the comment replies collection and track it
        const replies = firstComment.replies;
        replies.load("items");
        await context.sync();
        
        // Process the replies (e.g., log their content)
        console.log(`Found ${replies.items.length} replies`);
        
        // Untrack the collection to release memory
        replies.untrack();
        await context.sync();
    }
});
```

---

## Source

- https://learn.microsoft.com/en-us/javascript/api/word/word.commentreplycollection
