# Bookmark

**Package:** `word`

**API Set:** WordApi BETA PREVIEW ONLY

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents a single bookmark in a document, selection, or range. The `Bookmark` object is a member of the `Bookmark` collection. The [Word.BookmarkCollection](/en-us/javascript/api/word/word.bookmarkcollection) includes all the bookmarks listed in the Bookmark dialog box (Insert menu).

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a bookmark to load and read its properties

```typescript
await Word.run(async (context) => {
    // Get a bookmark by name
    const bookmark = context.document.getBookmarkByName("MyBookmark");
    
    // Access the context property to use it for loading properties
    const requestContext = bookmark.context;
    
    // Use the context to load bookmark properties
    bookmark.load("name,range/text");
    await requestContext.sync();
    
    console.log(`Bookmark name: ${bookmark.name}`);
    console.log(`Bookmark text: ${bookmark.range.text}`);
});
```

---

### end

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the ending character position of the bookmark.

#### Examples

**Example**: Get the ending character position of a bookmark named "Introduction" and display it in the console.

```typescript
await Word.run(async (context) => {
    const bookmark = context.document.getBookmarkByName("Introduction");
    bookmark.load("end");
    
    await context.sync();
    
    console.log(`The bookmark ends at character position: ${bookmark.end}`);
});
```

---

### isColumn

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns `true` if the bookmark is a table column.

#### Examples

**Example**: Check if a bookmark named "MyBookmark" is a table column and display an appropriate message based on the result.

```typescript
await Word.run(async (context) => {
    const bookmark = context.document.getBookmarkByName("MyBookmark");
    bookmark.load("isColumn");
    
    await context.sync();
    
    if (bookmark.isColumn) {
        console.log("The bookmark represents a table column.");
    } else {
        console.log("The bookmark is not a table column.");
    }
});
```

---

### isEmpty

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns `true` if the bookmark is empty.

#### Examples

**Example**: Check if a bookmark named "Section1" is empty and display an alert message based on the result.

```typescript
await Word.run(async (context) => {
    const bookmark = context.document.getBookmarkByName("Section1");
    bookmark.load("isEmpty");
    
    await context.sync();
    
    if (bookmark.isEmpty) {
        console.log("The bookmark 'Section1' is empty (contains no content).");
    } else {
        console.log("The bookmark 'Section1' contains content.");
    }
});
```

---

### name

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns the name of the `Bookmark` object.

#### Examples

**Example**: Get the name of the first bookmark in the document and display it in the console.

```typescript
await Word.run(async (context) => {
    const bookmarks = context.document.body.bookmarks;
    bookmarks.load("items");
    
    await context.sync();
    
    if (bookmarks.items.length > 0) {
        const firstBookmark = bookmarks.items[0];
        firstBookmark.load("name");
        
        await context.sync();
        
        console.log("Bookmark name: " + firstBookmark.name);
    } else {
        console.log("No bookmarks found in the document.");
    }
});
```

---

### range

**Type:** `Word.Range`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns a `Range` object that represents the portion of the document that's contained in the `Bookmark` object.

#### Examples

**Example**: Get the text content from a bookmark named "Introduction" and display it in the console.

```typescript
await Word.run(async (context) => {
    // Get the bookmark by name
    const bookmark = context.document.bookmarks.getByName("Introduction");
    
    // Get the range of the bookmark
    const bookmarkRange = bookmark.range;
    
    // Load the text property of the range
    bookmarkRange.load("text");
    
    await context.sync();
    
    // Display the bookmark's text content
    console.log("Bookmark text: " + bookmarkRange.text);
});
```

---

### start

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the starting character position of the bookmark.

#### Examples

**Example**: Get the starting character position of a bookmark named "Introduction" and display it in the console.

```typescript
await Word.run(async (context) => {
    const bookmark = context.document.getBookmarkByName("Introduction");
    bookmark.load("start");
    
    await context.sync();
    
    console.log(`The bookmark starts at character position: ${bookmark.start}`);
});
```

---

### storyType

**Type:** `Word.StoryType | "MainText" | "Footnotes" | "Endnotes" | "Comments" | "TextFrame" | "EvenPagesHeader" | "PrimaryHeader" | "EvenPagesFooter" | "PrimaryFooter" | "FirstPageHeader" | "FirstPageFooter" | "FootnoteSeparator" | "FootnoteContinuationSeparator" | "FootnoteContinuationNotice" | "EndnoteSeparator" | "EndnoteContinuationSeparator" | "EndnoteContinuationNotice"`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns the story type for the bookmark.

#### Examples

**Example**: Get the story type of a bookmark to determine where in the document it is located (e.g., main text, header, footer, footnotes)

```typescript
await Word.run(async (context) => {
    // Get a bookmark by name
    const bookmark = context.document.getBookmarkByName("MyBookmark");
    
    // Load the storyType property
    bookmark.load("storyType");
    
    await context.sync();
    
    // Display the story type
    console.log(`Bookmark is located in: ${bookmark.storyType}`);
    
    // You can use the story type to determine bookmark location
    if (bookmark.storyType === Word.StoryType.mainText) {
        console.log("Bookmark is in the main document body");
    } else if (bookmark.storyType === Word.StoryType.primaryHeader) {
        console.log("Bookmark is in the primary header");
    } else if (bookmark.storyType === Word.StoryType.footnotes) {
        console.log("Bookmark is in the footnotes section");
    }
});
```

---

## Methods

### copyTo

**Kind:** `create`

Copies this bookmark to the new bookmark specified in the `name` argument and returns a `Bookmark` object.

#### Signature

**Parameters:**
- `name`: `string` (required)
  The name of the new bookmark.

**Returns:** `Word.Bookmark`

#### Examples

**Example**: Copy an existing bookmark named "Introduction" to create a new bookmark named "IntroductionCopy" in the document.

```typescript
await Word.run(async (context) => {
    // Get the bookmark named "Introduction"
    const originalBookmark = context.document.body.bookmarks.getByName("Introduction");
    
    // Copy the bookmark to a new bookmark named "IntroductionCopy"
    const copiedBookmark = originalBookmark.copyTo("IntroductionCopy");
    
    // Load the name property to verify the copy
    copiedBookmark.load("name");
    
    await context.sync();
    
    console.log(`Bookmark copied successfully: ${copiedBookmark.name}`);
});
```

---

### delete

**Kind:** `delete`

Deletes the bookmark.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Delete a bookmark named "Section1" from the document

```typescript
await Word.run(async (context) => {
    // Get the bookmark by name
    const bookmark = context.document.getBookmarkByName("Section1");
    
    // Delete the bookmark
    bookmark.delete();
    
    await context.sync();
    console.log("Bookmark 'Section1' has been deleted.");
});
```

---

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.BookmarkLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.Bookmark`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.Bookmark`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.Bookmark`

#### Examples

**Example**: Load and display the name and text content of the first bookmark in the document

```typescript
await Word.run(async (context) => {
    // Get the first bookmark in the document
    const bookmarks = context.document.body.bookmarks;
    bookmarks.load("items");
    await context.sync();
    
    if (bookmarks.items.length > 0) {
        const firstBookmark = bookmarks.items[0];
        
        // Load specific properties of the bookmark
        firstBookmark.load("name, text");
        await context.sync();
        
        // Display the loaded properties
        console.log(`Bookmark name: ${firstBookmark.name}`);
        console.log(`Bookmark text: ${firstBookmark.text}`);
    } else {
        console.log("No bookmarks found in the document");
    }
});
```

---

### select

**Kind:** `read`

Selects the bookmark.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Select a bookmark named "Introduction" in the document to highlight its location for the user

```typescript
await Word.run(async (context) => {
    // Get the bookmark named "Introduction"
    const bookmark = context.document.bookmarks.getByName("Introduction");
    
    // Select the bookmark
    bookmark.select();
    
    await context.sync();
});
```

---

### set

**Kind:** `configure`

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `Interfaces.BookmarkUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.Bookmark` (required)

  **Returns:** `void`

#### Examples

**Example**: Update multiple properties of a bookmark at once, setting both its name and retrieving its text content

```typescript
await Word.run(async (context) => {
    // Get the first bookmark in the document
    const bookmarks = context.document.body.bookmarks;
    bookmarks.load("items");
    await context.sync();
    
    if (bookmarks.items.length > 0) {
        const bookmark = bookmarks.items[0];
        
        // Set multiple properties at once using the set() method
        bookmark.set({
            name: "UpdatedBookmarkName"
        });
        
        // Load properties to verify the changes
        bookmark.load("name, range/text");
        await context.sync();
        
        console.log(`Bookmark name: ${bookmark.name}`);
        console.log(`Bookmark text: ${bookmark.range.text}`);
    }
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.Bookmark` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.BookmarkData`) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.BookmarkData`

#### Examples

**Example**: Serialize a bookmark's properties to JSON format for logging or data transfer purposes

```typescript
await Word.run(async (context) => {
    // Get the first bookmark in the document
    const bookmarks = context.document.body.bookmarks;
    bookmarks.load("items");
    await context.sync();
    
    if (bookmarks.items.length > 0) {
        const bookmark = bookmarks.items[0];
        
        // Load properties we want to serialize
        bookmark.load("name, type, range/text");
        await context.sync();
        
        // Convert bookmark to plain JavaScript object
        const bookmarkData = bookmark.toJSON();
        
        // Now you can use the plain object for logging or data transfer
        console.log("Bookmark as JSON:", JSON.stringify(bookmarkData, null, 2));
        console.log("Bookmark name:", bookmarkData.name);
    }
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.Bookmark`

#### Examples

**Example**: Track a bookmark object across multiple sync calls to maintain its reference while modifying its properties and content outside of a single batch operation.

```typescript
await Word.run(async (context) => {
    // Get a bookmark by name
    const bookmark = context.document.getBookmarkByName("MyBookmark");
    bookmark.load("name,range/text");
    
    // Track the bookmark to use it across multiple sync calls
    bookmark.track();
    
    await context.sync();
    
    console.log(`Original bookmark text: ${bookmark.range.text}`);
    
    // Modify the bookmark's range in a subsequent operation
    bookmark.range.insertText(" - Updated", Word.InsertLocation.end);
    
    await context.sync();
    
    console.log(`Updated bookmark text: ${bookmark.range.text}`);
    
    // Untrack when done to free up memory
    bookmark.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

#### Signature

**Returns:** `Word.Bookmark`

#### Examples

**Example**: Load a bookmark, use its properties, then untrack it to free memory when done

```typescript
await Word.run(async (context) => {
    // Get a bookmark by name
    const bookmark = context.document.getBookmarkByName("MyBookmark");
    
    // Load properties to use them
    bookmark.load("name,range/text");
    await context.sync();
    
    // Use the bookmark
    console.log(`Bookmark: ${bookmark.name}, Text: ${bookmark.range.text}`);
    
    // Untrack the bookmark to release memory
    bookmark.untrack();
    await context.sync();
});
```

---

## Source

- /en-us/javascript/api/word
- /en-us/javascript/api/word/word.bookmarkcollection
- /en-us/javascript/api/office/officeextension.clientobject
- /en-us/javascript/api/word/word.requestcontext
- /en-us/javascript/api/word/word.range
- /en-us/javascript/api/word/word.storytype
- /en-us/javascript/api/word/word.bookmark
- /en-us/javascript/api/word/word.interfaces.bookmarkloadoptions
- /en-us/javascript/api/word/word.interfaces.bookmarkupdatedata
- /en-us/javascript/api/office/officeextension.updateoptions
- /en-us/javascript/api/word/word.interfaces.bookmarkdata
- /en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member
