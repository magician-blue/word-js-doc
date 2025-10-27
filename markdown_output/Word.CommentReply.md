# Word.CommentReply

**Package:** `word`

**API Set:** WordApi 1.4 None

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents a comment reply in the document.

## Class Examples

**Example**: Replies to the first active comment in the selected content.

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

### authorEmail

**Type:** `string`

**Since:** WordApi 1.4

Gets the email of the comment reply's author.

#### Examples

**Example**: Get and display the email address of the author who wrote the first reply to the first comment in the document.

```typescript
await Word.run(async (context) => {
    // Get the first comment in the document
    const firstComment = context.document.body.getComments().getFirst();
    
    // Get the first reply to that comment
    const firstReply = firstComment.replies.getFirst();
    
    // Load the authorEmail property
    firstReply.load("authorEmail");
    
    await context.sync();
    
    // Display the author's email
    console.log(`Reply author email: ${firstReply.authorEmail}`);
});
```

---

### authorName

**Type:** `string`

**Since:** WordApi 1.4

Gets the name of the comment reply's author.

#### Examples

**Example**: Get and display the author name of the first reply to the first comment in the document.

```typescript
await Word.run(async (context) => {
    // Get the first comment in the document
    const firstComment = context.document.body.getComments().getFirst();
    
    // Get the first reply to that comment
    const firstReply = firstComment.replies.getFirst();
    
    // Load the authorName property
    firstReply.load("authorName");
    
    await context.sync();
    
    // Display the author name
    console.log(`Reply author: ${firstReply.authorName}`);
});
```

---

### content

**Type:** `string`

**Since:** WordApi 1.4

Specifies the comment reply's content. The string is plain text.

#### Examples

**Example**: Update the content of the first reply to the first comment in the document to say "I agree with this suggestion"

```typescript
await Word.run(async (context) => {
    const firstComment = context.document.body.getComments().getFirst();
    const firstReply = firstComment.replies.getFirst();
    
    firstReply.content = "I agree with this suggestion";
    
    await context.sync();
});
```

---

### contentRange

**Type:** `Word.CommentContentRange`

**Since:** WordApi 1.4

Specifies the commentReply's content range.

#### Examples

**Example**: Highlight the content range of the first reply in the first comment by applying a yellow background color to make it stand out.

```typescript
await Word.run(async (context) => {
    // Get the first comment in the document
    const comments = context.document.getCommentCollection();
    comments.load("items");
    await context.sync();
    
    const firstComment = comments.items[0];
    
    // Get the first reply of the comment
    const replies = firstComment.replies;
    replies.load("items");
    await context.sync();
    
    const firstReply = replies.items[0];
    
    // Access the content range of the reply
    const replyContentRange = firstReply.contentRange;
    replyContentRange.load("text");
    
    // Apply yellow highlight to the reply's content range
    replyContentRange.font.highlightColor = "yellow";
    
    await context.sync();
});
```

---

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a comment reply to load and read its content property, demonstrating how the context connects the add-in to the Office host application.

```typescript
await Word.run(async (context) => {
    // Get the first comment in the document
    const comments = context.document.body.getComments();
    comments.load("items");
    await context.sync();
    
    if (comments.items.length > 0) {
        // Get the first reply of the first comment
        const replies = comments.items[0].replies;
        replies.load("items");
        await context.sync();
        
        if (replies.items.length > 0) {
            const reply = replies.items[0];
            
            // Access the context property to use it for loading data
            const replyContext = reply.context;
            
            // Use the context to load the reply's content
            reply.load("content");
            await replyContext.sync();
            
            console.log("Reply content: " + reply.content);
        }
    }
});
```

---

### creationDate

**Type:** `Date`

**Since:** WordApi 1.4

Gets the creation date of the comment reply.

#### Examples

**Example**: Display the creation date of the first reply to the first comment in the document

```typescript
await Word.run(async (context) => {
    const firstComment = context.document.body.getComments().getFirst();
    const firstReply = firstComment.replies.getFirst();
    firstReply.load("creationDate");
    
    await context.sync();
    
    console.log("Reply created on: " + firstReply.creationDate.toLocaleDateString());
});
```

---

### id

**Type:** `string`

**Since:** WordApi 1.4

Gets the ID of the comment reply.

#### Examples

**Example**: Get and display the ID of the first reply to the first comment in the document

```typescript
await Word.run(async (context) => {
    // Get the first comment in the document
    const firstComment = context.document.body.getComments().getFirst();
    
    // Get the first reply to that comment
    const firstReply = firstComment.replies.getFirst();
    
    // Load the ID property
    firstReply.load("id");
    
    await context.sync();
    
    // Display the comment reply ID
    console.log("Comment Reply ID: " + firstReply.id);
});
```

---

### parentComment

**Type:** `Word.Comment`

**Since:** WordApi 1.4

Gets the parent comment of this reply.

#### Examples

**Example**: Get the parent comment of a reply and display its content in the console.

```typescript
await Word.run(async (context) => {
    // Get the first comment reply in the document
    const commentReplies = context.document.body.getComments().getFirstOrNullObject().replies;
    const firstReply = commentReplies.getFirstOrNullObject();
    
    // Get the parent comment of this reply
    const parentComment = firstReply.parentComment;
    
    // Load the content of the parent comment
    parentComment.load("content");
    
    await context.sync();
    
    if (!firstReply.isNullObject) {
        console.log("Parent comment content: " + parentComment.content);
    }
});
```

---

## Methods

### delete

**Kind:** `delete`

Deletes the comment reply.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Delete the first reply from the first comment in the document

```typescript
await Word.run(async (context) => {
    // Get the first comment in the document
    const firstComment = context.document.body.getComments().getFirst();
    
    // Get the first reply of that comment
    const firstReply = firstComment.replies.getFirst();
    
    // Delete the reply
    firstReply.delete();
    
    await context.sync();
});
```

---

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.CommentReplyLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.CommentReply`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.CommentReply`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.CommentReply`

#### Examples

**Example**: Load and display the content and author name of the first reply to the first comment in the document.

```typescript
await Word.run(async (context) => {
    // Get the first comment in the document
    const comments = context.document.body.getComments();
    const firstComment = comments.getFirst();
    
    // Get the first reply of the comment
    const replies = firstComment.replies;
    const firstReply = replies.getFirst();
    
    // Load specific properties of the reply
    firstReply.load("content, authorName");
    
    // Sync to execute the load command
    await context.sync();
    
    // Now we can read the loaded properties
    console.log("Reply Author: " + firstReply.authorName);
    console.log("Reply Content: " + firstReply.content);
});
```

---

### set

**Kind:** `write`

Sets multiple properties of an object at the same time.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `Interfaces.CommentReplyUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.CommentReply` (required)

  **Returns:** `void`

#### Examples

**Example**: Update multiple properties of an existing comment reply, setting both its content text and marking it as resolved

```typescript
await Word.run(async (context) => {
    // Get the first comment in the document
    const comments = context.document.body.getComments();
    const firstComment = comments.getFirstOrNullObject();
    
    // Get the first reply of that comment
    const replies = firstComment.replies;
    replies.load("items");
    await context.sync();
    
    const firstReply = replies.items[0];
    
    // Set multiple properties at once using set()
    firstReply.set({
        content: "Updated reply text with additional information",
        resolved: true
    });
    
    await context.sync();
    console.log("Comment reply properties updated successfully");
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.CommentReply` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.CommentReplyData`) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.CommentReplyData`

#### Examples

**Example**: Serialize a comment reply to JSON format for logging or data export purposes

```typescript
await Word.run(async (context) => {
    // Get the first comment in the document
    const comments = context.document.body.getComments();
    const firstComment = comments.getFirstOrNullObject();
    
    // Get the first reply from that comment
    const replies = firstComment.replies;
    const firstReply = replies.getFirstOrNullObject();
    
    // Load properties we want to serialize
    firstReply.load("content, authorName, creationDate");
    
    await context.sync();
    
    // Convert the CommentReply object to a plain JavaScript object
    const replyData = firstReply.toJSON();
    
    // Now you can use the plain object for logging, storage, or transmission
    console.log("Reply as JSON:", JSON.stringify(replyData, null, 2));
    console.log("Author:", replyData.authorName);
    console.log("Content:", replyData.content);
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.CommentReply`

#### Examples

**Example**: Track a comment reply object to maintain its reference across multiple sync calls when modifying its properties in separate operations.

```typescript
await Word.run(async (context) => {
    // Get the first comment in the document
    const comments = context.document.body.getComments();
    const firstComment = comments.getFirstOrNullObject();
    firstComment.load("replies");
    
    await context.sync();
    
    if (!firstComment.isNullObject && firstComment.replies.items.length > 0) {
        const reply = firstComment.replies.getFirst();
        
        // Track the reply object to use it across multiple sync calls
        reply.track();
        
        reply.load("content");
        await context.sync();
        
        // Now we can safely modify the reply in a subsequent operation
        reply.content = "Updated reply content";
        await context.sync();
        
        // Untrack when done to free up memory
        reply.untrack();
    }
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

#### Signature

**Returns:** `Word.CommentReply`

#### Examples

**Example**: Reply to a comment, then untrack the reply object to free up memory after the operation is complete.

```typescript
await Word.run(async (context) => {
    // Get the first comment in the document
    const comments = context.document.body.getComments();
    const firstComment = comments.getFirstOrNullObject();
    firstComment.load("replies");
    
    await context.sync();
    
    if (!firstComment.isNullObject) {
        // Add a reply to the comment
        const reply = firstComment.addReply("Thank you for your feedback!");
        
        // Track the reply object for changes
        reply.track();
        reply.load("content");
        
        await context.sync();
        
        console.log("Reply added: " + reply.content);
        
        // Untrack the reply to release memory
        reply.untrack();
        
        await context.sync();
    }
});
```

---

## Source

- https://learn.microsoft.com/en-us/javascript/api/word
