# Word.List

**Package:** `word`

**API Set:** WordApi 1.3

**Extends:** `OfficeExtension.ClientObject`

## Description

Contains a collection of [Word.Paragraph](/en-us/javascript/api/word/word.paragraph) objects.

## Class Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/20-lists/insert-list.yaml

// This example starts a new list with the second paragraph.
await Word.run(async (context) => {
  const paragraphs: Word.ParagraphCollection = context.document.body.paragraphs;
  paragraphs.load("$none");

  await context.sync();

  // Start new list using the second paragraph.
  const list: Word.List = paragraphs.items[1].startNewList();
  list.load("$none");

  await context.sync();

  // To add new items to the list, use Start or End on the insertLocation parameter.
  list.insertParagraph("New list item at the start of the list", "Start");
  const paragraph: Word.Paragraph = list.insertParagraph("New list item at the end of the list (set to list level 5)", "End");

  // Set up list level for the list item.
  paragraph.listItem.level = 4;

  // To add paragraphs outside the list, use Before or After.
  list.insertParagraph("New paragraph goes after (not part of the list)", "After");

  await context.sync();
});
```

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a List object to verify the connection between the add-in and Word, then use it to load and log list properties.

```typescript
await Word.run(async (context) => {
    const lists = context.document.body.lists;
    lists.load("items");
    await context.sync();
    
    if (lists.items.length > 0) {
        const firstList = lists.items[0];
        
        // Access the request context associated with the list object
        const listContext = firstList.context;
        
        // Use the context to load list properties
        firstList.load("id, levelTypes");
        await listContext.sync();
        
        console.log(`List ID: ${firstList.id}`);
        console.log(`List level types: ${firstList.levelTypes}`);
    }
});
```

---

### id

**Type:** `number`

**Since:** WordApi 1.3

Gets the list's id.

#### Examples

**Example**: Retrieve and display the ID of the first list in the document to identify it for future operations

```typescript
await Word.run(async (context) => {
    // Get the first list in the document
    const lists = context.document.body.lists;
    lists.load("items");
    await context.sync();
    
    if (lists.items.length > 0) {
        const firstList = lists.items[0];
        firstList.load("id");
        await context.sync();
        
        // Display the list's ID
        console.log(`List ID: ${firstList.id}`);
    } else {
        console.log("No lists found in the document");
    }
});
```

---

### levelExistences

**Type:** `boolean[]`

**Since:** WordApi 1.3

Checks whether each of the 9 levels exists in the list. A true value indicates the level exists, which means there is at least one list item at that level.

#### Examples

**Example**: Retrieve and display the level types and level existences information for the first list in the document.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/20-lists/organize-list.yaml

// Gets information about the first list in the document.
await Word.run(async (context) => {
  const lists: Word.ListCollection = context.document.body.lists;
  lists.load("items");

  await context.sync();

  if (lists.items.length === 0) {
    console.warn("There are no lists in this document.");
    return;
  }
  
  // Get the first list.
  const list: Word.List = lists.getFirst();
  list.load("levelTypes,levelExistences");

  await context.sync();

  const levelTypes  = list.levelTypes;
  console.log("Level types of the first list:");
  for (let i = 0; i < levelTypes.length; i++) {
    console.log(`- Level ${i + 1} (index ${i}): ${levelTypes[i]}`);
  }

  const levelExistences = list.levelExistences;
  console.log("Level existences of the first list:");
  for (let i = 0; i < levelExistences.length; i++) {
    console.log(`- Level ${i + 1} (index ${i}): ${levelExistences[i]}`);
  }
});
```

---

### levelTypes

**Type:** `Word.ListLevelType[]`

**Since:** WordApi 1.3

Gets all 9 level types in the list. Each type can be 'Bullet', 'Number', or 'Picture'.

#### Examples

**Example**: Retrieve and display the level types and level existences information for the first list in the document.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/20-lists/organize-list.yaml

// Gets information about the first list in the document.
await Word.run(async (context) => {
  const lists: Word.ListCollection = context.document.body.lists;
  lists.load("items");

  await context.sync();

  if (lists.items.length === 0) {
    console.warn("There are no lists in this document.");
    return;
  }
  
  // Get the first list.
  const list: Word.List = lists.getFirst();
  list.load("levelTypes,levelExistences");

  await context.sync();

  const levelTypes  = list.levelTypes;
  console.log("Level types of the first list:");
  for (let i = 0; i < levelTypes.length; i++) {
    console.log(`- Level ${i + 1} (index ${i}): ${levelTypes[i]}`);
  }

  const levelExistences = list.levelExistences;
  console.log("Level existences of the first list:");
  for (let i = 0; i < levelExistences.length; i++) {
    console.log(`- Level ${i + 1} (index ${i}): ${levelExistences[i]}`);
  }
});
```

---

### paragraphs

**Type:** `Word.ParagraphCollection`

**Since:** WordApi 1.3

Gets paragraphs in the list.

#### Examples

**Example**: Get all paragraphs from the first list in the document and log their text content to the console.

```typescript
await Word.run(async (context) => {
    const lists = context.document.body.lists;
    lists.load("items");
    await context.sync();
    
    if (lists.items.length > 0) {
        const firstList = lists.items[0];
        const paragraphs = firstList.paragraphs;
        paragraphs.load("text");
        await context.sync();
        
        paragraphs.items.forEach((paragraph, index) => {
            console.log(`Paragraph ${index + 1}: ${paragraph.text}`);
        });
    }
});
```

---

## Methods

### getLevelFont

**Kind:** `read`

Gets the font of the bullet, number, or picture at the specified level in the list.

#### Signature

**Parameters:**
- `level`: `number` (required)
  The level in the list.

**Returns:** `Word.Font`

#### Examples

**Example**: Get the font of the list items at level 0 and change its color to red

```typescript
await Word.run(async (context) => {
    // Get the first list in the document
    const lists = context.document.body.lists;
    lists.load("items");
    await context.sync();
    
    if (lists.items.length > 0) {
        const firstList = lists.items[0];
        
        // Get the font at level 0 (first level)
        const levelFont = firstList.getLevelFont(0);
        levelFont.color = "red";
        
        await context.sync();
    }
});
```

---

### getLevelParagraphs

**Kind:** `read`

Gets the paragraphs that occur at the specified level in the list.

#### Signature

**Parameters:**
- `level`: `number` (required)
  The level in the list.

**Returns:** `Word.ParagraphCollection`

#### Examples

**Example**: Get all paragraphs at level 1 in the first list and highlight them in yellow

```typescript
await Word.run(async (context) => {
    // Get the first list in the document
    const list = context.document.body.lists.getFirst();
    
    // Get all paragraphs at level 1
    const level1Paragraphs = list.getLevelParagraphs(1);
    
    // Highlight them in yellow
    level1Paragraphs.load("items");
    await context.sync();
    
    level1Paragraphs.items.forEach(paragraph => {
        paragraph.font.highlightColor = "yellow";
    });
    
    await context.sync();
});
```

---

### getLevelPicture

**Kind:** `read`

Gets the Base64-encoded string representation of the picture at the specified level in the list.

#### Signature

**Parameters:**
- `level`: `number` (required)
  The level in the list.

**Returns:** `OfficeExtension.ClientResult<string>`

#### Examples

**Example**: Get the picture used as a bullet for level 1 items in the first list and display it in the console

```typescript
await Word.run(async (context) => {
    // Get the first list in the document
    const lists = context.document.body.lists;
    lists.load("items");
    await context.sync();
    
    if (lists.items.length > 0) {
        const firstList = lists.items[0];
        
        // Get the Base64-encoded picture for level 1
        const pictureBase64 = firstList.getLevelPicture(1);
        await context.sync();
        
        console.log("Level 1 picture (Base64):", pictureBase64.value);
    } else {
        console.log("No lists found in the document");
    }
});
```

---

### getLevelString

**Kind:** `read`

Gets the bullet, number, or picture at the specified level as a string.

#### Signature

**Parameters:**
- `level`: `number` (required)
  The level in the list.

**Returns:** `OfficeExtension.ClientResult<string>`

#### Examples

**Example**: Get the numbering format string for level 1 of the first list in the document and display it in the console.

```typescript
await Word.run(async (context) => {
    // Get the first list in the document
    const lists = context.document.body.lists;
    lists.load("items");
    await context.sync();
    
    if (lists.items.length > 0) {
        const firstList = lists.items[0];
        
        // Get the level string for level 1 (0-indexed)
        const levelString = firstList.getLevelString(1);
        await context.sync();
        
        console.log("Level 1 format: " + levelString.value);
    }
});
```

---

### insertParagraph

**Kind:** `create`

Inserts a paragraph at the specified location.

#### Signature

**Parameters:**
- `paragraphText`: `string` (required)
  The paragraph text to be inserted.
- `insertLocation`: `Word.InsertLocation.start | Word.InsertLocation.end | Word.InsertLocation.before | Word.InsertLocation.after | "Start" | "End" | "Before" | "After"` (required)
  The value must be 'Start', 'End', 'Before', or 'After'.

**Returns:** `Word.Paragraph`

#### Examples

**Example**: Create a new list from the second paragraph, add list items at the start and end with different list levels, and insert a paragraph after the list that is not part of it.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/20-lists/insert-list.yaml

// This example starts a new list with the second paragraph.
await Word.run(async (context) => {
  const paragraphs: Word.ParagraphCollection = context.document.body.paragraphs;
  paragraphs.load("$none");

  await context.sync();

  // Start new list using the second paragraph.
  const list: Word.List = paragraphs.items[1].startNewList();
  list.load("$none");

  await context.sync();

  // To add new items to the list, use Start or End on the insertLocation parameter.
  list.insertParagraph("New list item at the start of the list", "Start");
  const paragraph: Word.Paragraph = list.insertParagraph("New list item at the end of the list (set to list level 5)", "End");

  // Set up list level for the list item.
  paragraph.listItem.level = 4;

  // To add paragraphs outside the list, use Before or After.
  list.insertParagraph("New paragraph goes after (not part of the list)", "After");

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
  - `options`: `Word.Interfaces.ListLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.List`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.List`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string }` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.List`

#### Examples

**Example**: Load and display the font name of each paragraph in the first list of the document

```typescript
await Word.run(async (context) => {
    // Get the first list in the document
    const lists = context.document.body.lists;
    lists.load("items");
    await context.sync();
    
    const firstList = lists.items[0];
    
    // Load the paragraphs collection and their font names
    firstList.paragraphs.load("font/name");
    await context.sync();
    
    // Display the font name of each paragraph in the list
    firstList.paragraphs.items.forEach((paragraph, index) => {
        console.log(`Paragraph ${index + 1} font: ${paragraph.font.name}`);
    });
});
```

---

### resetLevelFont

**Kind:** `configure`

Resets the font of the bullet, number, or picture at the specified level in the list.

#### Signature

**Parameters:**
- `level`: `number` (required)
  The level in the list.
- `resetFontName`: `boolean` (optional)
  Indicates whether to reset the font name. Default is false that indicates the font name is kept unchanged.

**Returns:** `void`

#### Examples

**Example**: Reset the font formatting of level 1 list items to the default font in the first list of the document

```typescript
await Word.run(async (context) => {
    // Get the first list in the document
    const list = context.document.body.lists.getFirst();
    
    // Reset the font for level 1 items (level is 0-indexed)
    // The second parameter 'true' indicates to reset the font name
    list.resetLevelFont(0, true);
    
    await context.sync();
    
    console.log("Level 1 list font has been reset to default");
});
```

---

### setLevelAlignment

**Kind:** `configure`

Sets the alignment of the bullet, number, or picture at the specified level in the list.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `level`: `number` (required)
    The level in the list.
  - `alignment`: `Word.Alignment` (required)
    The level alignment that must be 'Left', 'Centered', or 'Right'.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `level`: `number` (required)
    The level in the list.
  - `alignment`: `"Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified"` (required)
    The level alignment that must be 'Left', 'Centered', or 'Right'.

  **Returns:** `void`

#### Examples

**Example**: Set the alignment of level 1 list items to center and level 2 list items to right in the first list of the document

```typescript
await Word.run(async (context) => {
    // Get the first list in the document
    const list = context.document.body.lists.getFirst();
    
    // Set level 1 alignment to center
    list.setLevelAlignment(1, Word.Alignment.centered);
    
    // Set level 2 alignment to right
    list.setLevelAlignment(2, Word.Alignment.right);
    
    await context.sync();
});
```

---

### setLevelBullet

**Kind:** `configure`

Sets the bullet format at the specified level in the list. If the bullet is 'Custom', the charCode is required.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `level`: `number` (required)
    The level in the list.
  - `listBullet`: `Word.ListBullet` (required)
    The bullet.
  - `charCode`: `number` (optional)
    The bullet character's code value. Used only if the bullet is 'Custom'.
  - `fontName`: `string` (optional)
    The bullet's font name. Used only if the bullet is 'Custom'.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `level`: `number` (required)
    The level in the list.
  - `listBullet`: `"Custom" | "Solid" | "Hollow" | "Square" | "Diamonds" | "Arrow" | "Checkmark"` (required)
    The bullet.
  - `charCode`: `number` (optional)
    The bullet character's code value. Used only if the bullet is 'Custom'.
  - `fontName`: `string` (optional)
    The bullet's font name. Used only if the bullet is 'Custom'.

  **Returns:** `void`

#### Examples

**Example**: Set the bullet type to arrow for list level 5 in a Word document list.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/20-lists/organize-list.yaml

// Inserts a list starting with the first paragraph then set numbering and bullet types of the list items.
await Word.run(async (context) => {
  const paragraphs: Word.ParagraphCollection = context.document.body.paragraphs;
  paragraphs.load("$none");

  await context.sync();

  // Use the first paragraph to start a new list.
  const list: Word.List = paragraphs.items[0].startNewList();
  list.load("$none");

  await context.sync();

  // To add new items to the list, use Start or End on the insertLocation parameter.
  list.insertParagraph("New list item at the start of the list", "Start");
  const paragraph: Word.Paragraph = list.insertParagraph("New list item at the end of the list (set to list level 5)", "End");

  // Set numbering for list level 1.
  list.setLevelNumbering(0, Word.ListNumbering.arabic);

  // Set bullet type for list level 5.
  list.setLevelBullet(4, Word.ListBullet.arrow);

  // Set list level for the last item in this list.
  paragraph.listItem.level = 4;

  list.load("levelTypes");

  await context.sync();
});
```

---

### setLevelIndents

**Kind:** `configure`

Sets the two indents of the specified level in the list.

#### Signature

**Parameters:**
- `level`: `number` (required)
  The level in the list.
- `textIndent`: `number` (required)
  The text indent in points. It is the same as paragraph left indent.
- `bulletNumberPictureIndent`: `number` (required)
  The relative indent, in points, of the bullet, number, or picture. It is the same as paragraph first line indent.

**Returns:** `void`

#### Examples

**Example**: Set the indents for level 1 of the first list in the document, with a text indent of 0.5 inches and a bullet/number indent of 0.25 inches

```typescript
await Word.run(async (context) => {
    // Get the first list in the document
    const list = context.document.body.lists.getFirst();
    
    // Set indents for level 1 (0-indexed)
    // textIndent: 0.5 inches (36 points)
    // bulletNumberPictureIndent: 0.25 inches (18 points)
    list.setLevelIndents(0, 36, 18);
    
    await context.sync();
});
```

---

### setLevelNumbering

**Kind:** `configure`

Sets the numbering format at the specified level in the list.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `level`: `number` (required)
    The level in the list.
  - `listNumbering`: `Word.ListNumbering` (required)
    The ordinal format.
  - `formatString`: `Array<string | number>` (optional)
    The numbering string format defined as an array of strings and/or integers. Each integer is a level of number type that is higher than or equal to this level. For example, an array of ["(", level - 1, ".", level, ")"] can define the format of "(2.c)", where 2 is the parent's item number and c is this level's item number.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `level`: `number` (required)
    The level in the list.
  - `listNumbering`: `"None" | "Arabic" | "UpperRoman" | "LowerRoman" | "UpperLetter" | "LowerLetter"` (required)
    The ordinal format.
  - `formatString`: `Array<string | number>` (optional)
    The numbering string format defined as an array of strings and/or integers. Each integer is a level of number type that is higher than or equal to this level. For example, an array of ["(", level - 1, ".", level, ")"] can define the format of "(2.c)", where 2 is the parent's item number and c is this level's item number.

  **Returns:** `void`

#### Examples

**Example**: Configure a list with Arabic numbering for level 1, add new items at the start and end of the list, and set the last item to level 5 with an arrow bullet type.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/20-lists/organize-list.yaml

// Inserts a list starting with the first paragraph then set numbering and bullet types of the list items.
await Word.run(async (context) => {
  const paragraphs: Word.ParagraphCollection = context.document.body.paragraphs;
  paragraphs.load("$none");

  await context.sync();

  // Use the first paragraph to start a new list.
  const list: Word.List = paragraphs.items[0].startNewList();
  list.load("$none");

  await context.sync();

  // To add new items to the list, use Start or End on the insertLocation parameter.
  list.insertParagraph("New list item at the start of the list", "Start");
  const paragraph: Word.Paragraph = list.insertParagraph("New list item at the end of the list (set to list level 5)", "End");

  // Set numbering for list level 1.
  list.setLevelNumbering(0, Word.ListNumbering.arabic);

  // Set bullet type for list level 5.
  list.setLevelBullet(4, Word.ListBullet.arrow);

  // Set list level for the last item in this list.
  paragraph.listItem.level = 4;

  list.load("levelTypes");

  await context.sync();
});
```

---

### setLevelPicture

**Kind:** `configure`

Sets the picture at the specified level in the list.

#### Signature

**Parameters:**
- `level`: `number` (required)
  The level in the list.
- `base64EncodedImage`: `string` (optional)
  The Base64-encoded image to be set. If not given, the default picture is set.

**Returns:** `void`

#### Examples

**Example**: Set a custom image as the bullet picture for level 0 items in the first list of the document

```typescript
await Word.run(async (context) => {
    // Get the first list in the document
    const list = context.document.body.lists.getFirst();
    
    // Base64 encoded image string (example: a small red circle)
    const base64Image = "iVBORw0KGgoAAAANSUhEUgAAAAoAAAAKCAYAAACNMs+9AAAAFUlEQVR42mP8z8BQz0AEYBxVSF+FABJADveWkH6oAAAAAElFTkSuQmCC";
    
    // Set the picture for level 0 (first level) of the list
    list.setLevelPicture(0, base64Image);
    
    await context.sync();
});
```

---

### setLevelStartingNumber

**Kind:** `configure`

Sets the starting number at the specified level in the list. Default value is 1.

#### Signature

**Parameters:**
- `level`: `number` (required)
  The level in the list.
- `startingNumber`: `number` (required)
  The number to start with.

**Returns:** `void`

#### Examples

**Example**: Set the starting number of level 1 to 5 and level 2 to 10 in the first list of the document

```typescript
await Word.run(async (context) => {
    const firstList = context.document.body.lists.getFirst();
    
    // Set level 1 to start at 5
    firstList.setLevelStartingNumber(1, 5);
    
    // Set level 2 to start at 10
    firstList.setLevelStartingNumber(2, 10);
    
    await context.sync();
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.List object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.ListData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.ListData`

#### Examples

**Example**: Serialize a list object to JSON format to log or export its properties for debugging purposes.

```typescript
await Word.run(async (context) => {
    // Get the first list in the document
    const list = context.document.body.lists.getFirst();
    
    // Load properties we want to serialize
    list.load("id,levelTypes,levelExistences");
    
    await context.sync();
    
    // Convert the list object to a plain JavaScript object
    const listJson = list.toJSON();
    
    // Now you can use the plain object (e.g., log it, send it to a server, etc.)
    console.log("List as JSON:", JSON.stringify(listJson, null, 2));
    console.log("List ID:", listJson.id);
    console.log("Level types:", listJson.levelTypes);
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.List`

#### Examples

**Example**: Track a list object across multiple sync calls to maintain its reference while modifying list properties and paragraphs

```typescript
await Word.run(async (context) => {
    // Get the first list in the document
    const list = context.document.body.lists.getFirst();
    
    // Track the list object to maintain reference across sync calls
    list.track();
    
    // Load properties
    list.load("levelTypes");
    await context.sync();
    
    // Use the list in subsequent operations
    console.log("List level types:", list.levelTypes);
    
    // Modify list properties
    const firstParagraph = list.paragraphs.getFirst();
    firstParagraph.load("text");
    await context.sync();
    
    console.log("First item:", firstParagraph.text);
    
    // Untrack when done to free up memory
    list.untrack();
    
    await context.sync();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.List`

#### Examples

**Example**: Get a list from the document, use it to retrieve paragraph information, then release it from memory tracking to improve performance.

```typescript
await Word.run(async (context) => {
    // Get the first list in the document
    const list = context.document.body.lists.getFirst();
    
    // Load properties to work with the list
    list.load("id");
    
    await context.sync();
    
    // Use the list (e.g., log its ID)
    console.log("List ID: " + list.id);
    
    // Release the list from memory tracking when done
    list.untrack();
    
    await context.sync();
});
```

---

## Source

- https://docs.microsoft.com/en-us/javascript/api/word/word.list
