# Word.Comment

**Package:** `word`

**API Set:** WordApi 1.4 None

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents a comment in the document.

## Class Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-comments.yaml

// Sets a comment on the selected content.
await Word.run(async (context) => {
  const text = (document.getElementById("comment-text") as HTMLInputElement).value;
  const comment: Word.Comment = context.document.getSelection().insertComment(text);

  // Load object to log in the console.
  comment.load();
  await context.sync();

  console.log("Comment inserted:", comment);
});
```

## Properties

### authorEmail

**Type:** `None`

Gets the email of the comment's author.

#### Examples

**Example**: Display the email address of the author who created the first comment in the document

```typescript
await Word.run(async (context) => {
    const comments = context.document.body.getComments();
    comments.load("items");
    
    await context.sync();
    
    if (comments.items.length > 0) {
        const firstComment = comments.items[0];
        firstComment.load("authorEmail");
        
        await context.sync();
        
        console.log("Comment author email: " + firstComment.authorEmail);
    }
});
```

---

### authorName

**Type:** `None`

Gets the name of the comment's author.

#### Examples

**Example**: Get and display the author name of the first comment in the document

```typescript
await Word.run(async (context) => {
    const comments = context.document.body.getComments();
    comments.load("items");
    await context.sync();

    if (comments.items.length > 0) {
        const firstComment = comments.items[0];
        firstComment.load("authorName");
        await context.sync();

        console.log("Comment author: " + firstComment.authorName);
    }
});
```

---

### content

**Type:** `None`

Specifies the comment's content as plain text.

#### Examples

**Example**: Get the text content of the first comment in the document and display it in the console

```typescript
await Word.run(async (context) => {
    const comments = context.document.body.getComments();
    comments.load("items");
    await context.sync();

    if (comments.items.length > 0) {
        const firstComment = comments.items[0];
        firstComment.load("content");
        await context.sync();

        console.log("Comment content: " + firstComment.content);
    }
});
```

---

### contentRange

**Type:** `None`

Specifies the comment's content range.

#### Examples

**Example**: Highlight the content range of the first comment in the document with yellow color to visually identify the commented text.

```typescript
await Word.run(async (context) => {
    // Get the first comment in the document
    const comments = context.document.body.getComments();
    const firstComment = comments.getFirst();
    
    // Get the content range of the comment
    const contentRange = firstComment.contentRange;
    
    // Highlight the content range with yellow
    contentRange.font.highlightColor = "yellow";
    
    // Load and sync to apply changes
    await context.sync();
    
    console.log("Comment content range highlighted");
});
```

---

### context

**Type:** `None`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access a comment's context to verify the connection to the Office host application and log its debug information

```typescript
await Word.run(async (context) => {
    // Get the first comment in the document
    const comments = context.document.body.getComments();
    comments.load("items");
    await context.sync();
    
    if (comments.items.length > 0) {
        const firstComment = comments.items[0];
        
        // Access the comment's context property
        const commentContext = firstComment.context;
        
        // Use the context to perform operations
        // For example, check if debug mode is enabled
        console.log("Context debug info:", commentContext.debugInfo);
        
        // The context connects the add-in to the Office host
        // and is used internally for all API calls
        firstComment.load("content");
        await commentContext.sync();
        
        console.log("Comment content:", firstComment.content);
    }
});
```

---

### creationDate

**Type:** `None`

Gets the creation date of the comment.

#### Examples

**Example**: Display the creation date of the first comment in the document

```typescript
await Word.run(async (context) => {
    const comments = context.document.body.getComments();
    comments.load("items");
    await context.sync();

    if (comments.items.length > 0) {
        const firstComment = comments.items[0];
        firstComment.load("creationDate");
        await context.sync();

        console.log("Comment created on: " + firstComment.creationDate);
    }
});
```

---

### id

**Type:** `None`

Gets the ID of the comment.

#### Examples

**Example**: Get the ID of the first comment in the document and display it in the console.

```typescript
await Word.run(async (context) => {
    const comments = context.document.body.getComments();
    comments.load("items");
    
    await context.sync();
    
    if (comments.items.length > 0) {
        const firstComment = comments.items[0];
        firstComment.load("id");
        
        await context.sync();
        
        console.log("Comment ID: " + firstComment.id);
    }
});
```

---

### replies

**Type:** `None`

Gets the collection of reply objects associated with the comment.

#### Examples

**Example**: Get all replies to the first comment in the document and display the count and content of each reply in the console.

```typescript
await Word.run(async (context) => {
    // Get the first comment in the document
    const comments = context.document.body.getComments();
    const firstComment = comments.getFirst();
    
    // Get the replies collection for this comment
    const replies = firstComment.replies;
    replies.load("items");
    
    await context.sync();
    
    // Display information about the replies
    console.log(`Total replies: ${replies.items.length}`);
    
    replies.items.forEach((reply, index) => {
        reply.load("content, authorName");
    });
    
    await context.sync();
    
    replies.items.forEach((reply, index) => {
        console.log(`Reply ${index + 1}: ${reply.content} (by ${reply.authorName})`);
    });
});
```

---

### resolved

**Type:** `None`

Specifies the comment thread's status. Setting to true resolves the comment thread. Getting a value of true means that the comment thread is resolved.

#### Examples

**Example**: Mark a comment thread as resolved after reviewing it

```typescript
await Word.run(async (context) => {
    // Get the first comment in the document
    const comment = context.document.body.getComments().getFirst();
    
    // Mark the comment as resolved
    comment.resolved = true;
    
    await context.sync();
    
    console.log("Comment has been marked as resolved");
});
```

---

## Methods

### delete

**Kind:** `delete`

Deletes the comment and its replies.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Delete all comments in the document that contain the word "outdated"

```typescript
await Word.run(async (context) => {
    // Get all comments in the document
    const comments = context.document.body.getComments();
    comments.load("items");
    
    await context.sync();
    
    // Loop through comments and delete those containing "outdated"
    for (let i = 0; i < comments.items.length; i++) {
        const comment = comments.items[i];
        comment.load("content");
        await context.sync();
        
        if (comment.content.toLowerCase().includes("outdated")) {
            comment.delete();
        }
    }
    
    await context.sync();
});
```

---

### getRange

**Kind:** `read`

Gets the range in the main document where the comment is on.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Highlight the text range that a comment is attached to by applying a yellow background color to help visualize the commented area.

```typescript
await Word.run(async (context) => {
    // Get the first comment in the document
    const comments = context.document.body.getComments();
    const firstComment = comments.getFirst();
    
    // Get the range where the comment is attached
    const commentRange = firstComment.getRange();
    
    // Highlight the commented range with yellow background
    commentRange.font.highlightColor = "yellow";
    
    await context.sync();
    
    console.log("Highlighted the range where the comment is located");
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

**Example**: Load and display the content and author name of the first comment in the document

```typescript
await Word.run(async (context) => {
    // Get the first comment in the document
    const comments = context.document.body.getComments();
    const firstComment = comments.getFirstOrNullObject();
    
    // Load specific properties of the comment
    firstComment.load("content, authorName, creationDate");
    
    await context.sync();
    
    if (!firstComment.isNullObject) {
        console.log(`Author: ${firstComment.authorName}`);
        console.log(`Content: ${firstComment.content}`);
        console.log(`Created: ${firstComment.creationDate}`);
    } else {
        console.log("No comments found in the document");
    }
});
```

---

### reply

**Kind:** `create`

Adds a new reply to the end of the comment thread.

#### Signature

**Parameters:**
- `replyText`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Add a reply "I've addressed this feedback" to the first comment in the document

```typescript
await Word.run(async (context) => {
    const comments = context.document.body.getComments();
    comments.load("items");
    
    await context.sync();
    
    if (comments.items.length > 0) {
        const firstComment = comments.items[0];
        firstComment.reply("I've addressed this feedback");
        
        await context.sync();
    }
});
```

---

### set

**Kind:** `configure`

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `None` (required)
  - `options`: `None` (required)

  **Returns:** `None`

**Overload 2:**

  **Parameters:**
  - `properties`: `None` (required)

  **Returns:** `None`

#### Examples

**Example**: Update an existing comment's properties to mark it as resolved and change its content text

```typescript
await Word.run(async (context) => {
    // Get the first comment in the document
    const comment = context.document.body.getComments().getFirst();
    
    // Set multiple properties at once using the set() method
    comment.set({
        resolved: true,
        content: "This issue has been addressed and resolved."
    });
    
    await context.sync();
    
    console.log("Comment properties updated successfully");
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.Comment object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.CommentData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Retrieve a comment from the document and serialize it to JSON format for logging or data export purposes.

```typescript
await Word.run(async (context) => {
    // Get the first comment in the document
    const comments = context.document.body.getComments();
    const firstComment = comments.getFirstOrNullObject();
    
    // Load properties we want to serialize
    firstComment.load("content,authorName,creationDate,resolved");
    
    await context.sync();
    
    if (!firstComment.isNullObject) {
        // Convert the comment to a plain JavaScript object
        const commentData = firstComment.toJSON();
        
        // Now you can use the plain object (e.g., log it, send to server, etc.)
        console.log("Comment as JSON:", JSON.stringify(commentData, null, 2));
        console.log("Author:", commentData.authorName);
        console.log("Content:", commentData.content);
    } else {
        console.log("No comments found in the document.");
    }
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent

#### Signature

**Returns:** `None`

#### Examples

**Example**: Get a comment from the document, track it to maintain its reference across multiple sync calls, then modify its content and resolved status in separate operations.

```typescript
await Word.run(async (context) => {
    // Get the first comment in the document
    const comments = context.document.body.getComments();
    comments.load("items");
    await context.sync();
    
    const firstComment = comments.items[0];
    
    // Track the comment object to use it across multiple sync calls
    firstComment.track();
    
    // First operation: load and display current content
    firstComment.load("content");
    await context.sync();
    console.log("Current comment: " + firstComment.content);
    
    // Second operation: modify the comment content
    firstComment.content = "Updated comment text";
    await context.sync();
    
    // Third operation: mark the comment as resolved
    firstComment.resolved = true;
    await context.sync();
    
    // Untrack when done to free up memory
    firstComment.untrack();
});
```

---

## Source

- https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-comments.yaml
