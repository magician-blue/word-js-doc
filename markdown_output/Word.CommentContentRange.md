# Word.CommentContentRange

**Package:** `word`

**API Set:** WordApi 1.4

**Extends:** `OfficeExtension.ClientObject`

## Class Examples

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

## Properties

### bold

**Type:** `boolean`

**Since:** 1.4

Specifies a value that indicates whether the comment text is bold.

#### Examples

**Example**: Make all text in a comment content range bold

```typescript
await Word.run(async (context) => {
    // Get the first comment in the document
    const comment = context.document.body.paragraphs.getFirst().getComments().getFirst();
    const commentContentRange = comment.contentRange;
    
    // Make the comment text bold
    commentContentRange.bold = true;
    
    await context.sync();
});
```

---

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a CommentContentRange object to verify the connection between the add-in and Word application before performing operations on comment content.

```typescript
await Word.run(async (context) => {
    // Get the first comment in the document
    const comments = context.document.getCommentCollection();
    comments.load("items");
    await context.sync();
    
    if (comments.items.length > 0) {
        const comment = comments.items[0];
        const contentRange = comment.getRange();
        
        // Access the request context from the CommentContentRange
        const requestContext = contentRange.context;
        
        // Use the context to verify connection and perform operations
        console.log("Context connected:", requestContext !== null);
        
        // Load and display the comment content using the context
        contentRange.load("text");
        await requestContext.sync();
        
        console.log("Comment content:", contentRange.text);
    }
});
```

---

### hyperlink

**Type:** `string`

**Since:** 1.4

Gets the first hyperlink in the range, or sets a hyperlink on the range. All hyperlinks in the range are deleted when you set a new hyperlink on the range.

#### Examples

**Example**: Set a hyperlink to "https://www.example.com" on the content range of a comment

```typescript
await Word.run(async (context) => {
    // Get the first comment in the document
    const comment = context.document.body.getComments().getFirst();
    const commentContentRange = comment.contentRange;
    
    // Set a hyperlink on the comment content range
    commentContentRange.hyperlink = "https://www.example.com";
    
    await context.sync();
    console.log("Hyperlink set on comment content range");
});
```

---

### isEmpty

**Type:** `boolean`

**Since:** 1.4

Checks whether the range length is zero.

#### Examples

**Example**: Check if a comment content range is empty and display an alert message based on the result

```typescript
await Word.run(async (context) => {
    const comment = context.document.body.paragraphs.getFirst().getComments().getFirstOrNullObject();
    await context.sync();
    
    if (!comment.isNullObject) {
        const contentRange = comment.contentRange;
        contentRange.load("isEmpty");
        await context.sync();
        
        if (contentRange.isEmpty) {
            console.log("The comment content range is empty.");
        } else {
            console.log("The comment content range contains content.");
        }
    }
});
```

---

### italic

**Type:** `boolean`

**Since:** 1.4

Specifies a value that indicates whether the comment text is italicized.

#### Examples

**Example**: Make all text in a comment's content range italic

```typescript
await Word.run(async (context) => {
    // Get the first comment in the document
    const comment = context.document.body.comments.getFirst();
    const contentRange = comment.contentRange;
    
    // Set the text to italic
    contentRange.italic = true;
    
    await context.sync();
});
```

---

### strikeThrough

**Type:** `boolean`

**Since:** 1.4

Specifies a value that indicates whether the comment text has a strikethrough.

#### Examples

**Example**: Apply strikethrough formatting to the text content of a comment in the document

```typescript
await Word.run(async (context) => {
    // Get the first comment in the document
    const comment = context.document.body.getComments().getFirst();
    
    // Get the comment's content range
    const commentContentRange = comment.contentRange;
    
    // Apply strikethrough to the comment text
    commentContentRange.strikeThrough = true;
    
    await context.sync();
});
```

---

### text

**Type:** `string`

**Since:** 1.4

Gets the text of the comment range.

#### Examples

**Example**: Read and display the text content from a comment range in the document

```typescript
await Word.run(async (context) => {
    // Get the first comment in the document
    const comments = context.document.body.getComments();
    comments.load("items");
    await context.sync();
    
    if (comments.items.length > 0) {
        const firstComment = comments.items[0];
        
        // Get the comment content range
        const commentContentRange = firstComment.getRange();
        commentContentRange.load("text");
        await context.sync();
        
        // Display the text from the comment range
        console.log("Comment range text: " + commentContentRange.text);
    }
});
```

---

### underline

**Type:** `Word.UnderlineType | "Mixed" | "None" | "Hidden" | "DotLine" | "Single" | "Word" | "Double" | "Thick" | "Dotted" | "DottedHeavy" | "DashLine" | "DashLineHeavy" | "DashLineLong" | "DashLineLongHeavy" | "DotDashLine" | "DotDashLineHeavy" | "TwoDotDashLine" | "TwoDotDashLineHeavy" | "Wave" | "WaveHeavy" | "WaveDouble"`

**Since:** 1.4

Specifies a value that indicates the comment text's underline type. 'None' if the comment text isn't underlined.

#### Examples

**Example**: Apply a double underline style to text within a comment content range

```typescript
await Word.run(async (context) => {
    // Get the first comment in the document
    const comment = context.document.body.getComments().getFirst();
    const commentContentRange = comment.contentRange;
    
    // Apply double underline to the comment text
    commentContentRange.underline = "Double";
    
    await context.sync();
});
```

---

## Methods

### insertText

**Kind:** `write`

Inserts text into at the specified location. Note: For the modern comment, the content range tracked across context turns to empty if any revision to the comment is posted through the UI.

#### Signature

**Parameters:**
- `text`: `string` (required)
  The text to be inserted in to the CommentContentRange.
- `insertLocation`: `Word.InsertLocation | "Replace" | "Start" | "End" | "Before" | "After"` (required)
  The value must be 'Replace', 'Start', 'End', 'Before', or 'After'.

**Returns:** `Word.CommentContentRange`

#### Examples

**Example**: Insert the text "Please review this section carefully." at the beginning of a comment's content range in a Word document.

```typescript
await Word.run(async (context) => {
    // Get the first comment in the document
    const comments = context.document.getCommentContentRanges();
    const firstComment = comments.getFirst();
    
    // Insert text at the beginning of the comment content range
    firstComment.insertText("Please review this section carefully.", Word.InsertLocation.start);
    
    await context.sync();
});
```

---

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.CommentContentRangeLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.CommentContentRange`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.CommentContentRange`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.CommentContentRange`

#### Examples

**Example**: Load and display the text content of a comment content range associated with a comment in the document.

```typescript
await Word.run(async (context) => {
    // Get the first comment in the document
    const comments = context.document.getCommentCollection();
    comments.load("items");
    await context.sync();
    
    if (comments.items.length > 0) {
        const comment = comments.items[0];
        
        // Get the content range of the comment
        const contentRange = comment.getRange();
        
        // Load the text property of the content range
        contentRange.load("text");
        await context.sync();
        
        // Display the loaded text
        console.log("Comment content range text: " + contentRange.text);
    }
});
```

---

### set

**Kind:** `write`

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `Interfaces.CommentContentRangeUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.CommentContentRange` (required)

  **Returns:** `void`

#### Examples

**Example**: Update multiple properties of a comment content range to highlight and format a specific portion of text within a comment

```typescript
await Word.run(async (context) => {
    // Get the first comment in the document
    const comments = context.document.getCommentContentRanges();
    const firstComment = comments.getFirst();
    
    // Set multiple properties on the comment content range at once
    firstComment.set({
        font: {
            bold: true,
            color: "blue",
            size: 12
        }
    });
    
    await context.sync();
    console.log("Comment content range properties updated successfully");
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.CommentContentRange object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.CommentContentRangeData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.CommentContentRangeData`

#### Examples

**Example**: Get a JSON representation of a comment content range to log or store its properties for debugging purposes.

```typescript
await Word.run(async (context) => {
    // Get the first comment in the document
    const comments = context.document.getCommentByIdOrNullObject("comment1");
    const comment = comments.getFirstOrNullObject();
    
    // Get the content range of the comment
    const contentRange = comment.contentRange;
    
    // Load properties we want to serialize
    contentRange.load("text, isEmpty");
    
    await context.sync();
    
    // Convert to JSON for logging or storage
    const jsonData = contentRange.toJSON();
    console.log("Comment content range data:", JSON.stringify(jsonData, null, 2));
    
    // The JSON object contains shallow copies of loaded properties
    console.log("Text:", jsonData.text);
    console.log("Is Empty:", jsonData.isEmpty);
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.CommentContentRange`

#### Examples

**Example**: Track a comment content range object to maintain its reference across multiple sync calls when highlighting the text within a comment's content range.

```typescript
await Word.run(async (context) => {
    // Get the first comment in the document
    const comments = context.document.getCommentCollection();
    comments.load("items");
    await context.sync();
    
    if (comments.items.length > 0) {
        const comment = comments.items[0];
        const contentRange = comment.getRange();
        
        // Track the content range to use it across multiple sync calls
        contentRange.track();
        
        // Load properties
        contentRange.load("text");
        await context.sync();
        
        console.log("Comment content: " + contentRange.text);
        
        // Use the tracked object in another operation after sync
        contentRange.font.highlightColor = "yellow";
        await context.sync();
        
        // Untrack when done
        contentRange.untrack();
    }
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.CommentContentRange`

#### Examples

**Example**: Load a comment content range, use its properties, then untrack it to free memory when done.

```typescript
await Word.run(async (context) => {
    // Get the first comment in the document
    const comments = context.document.body.getComments();
    comments.load("items");
    await context.sync();
    
    if (comments.items.length > 0) {
        const comment = comments.items[0];
        const contentRange = comment.getRange();
        
        // Track the object to use it across sync calls
        contentRange.track();
        contentRange.load("text");
        await context.sync();
        
        // Use the content range
        console.log("Comment content: " + contentRange.text);
        
        // Untrack when done to release memory
        contentRange.untrack();
        await context.sync();
    }
});
```

---

## Source

- https://docs.microsoft.com/en-us/javascript/api/word/word.commentcontentrange
