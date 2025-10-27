# Word.ListFormat

**Package:** `word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents the list formatting characteristics of a range.

## Properties

### context

**Type:** `RequestContext`

**Since:** WordApi BETA (PREVIEW ONLY)

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a ListFormat object to load and read list properties for a selected paragraph.

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.getSelection().paragraphs.getFirst();
    const listFormat = paragraph.listFormat;
    
    // Access the request context associated with the ListFormat object
    const requestContext = listFormat.context;
    
    // Use the context to load properties
    listFormat.load("listLevelNumber,listString");
    
    await requestContext.sync();
    
    console.log("List Level: " + listFormat.listLevelNumber);
    console.log("List String: " + listFormat.listString);
});
```

---

### isSingleList

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Indicates whether the `ListFormat` object contains a single list.

#### Examples

**Example**: Check if a selected range contains only a single list and display an alert with the result.

```typescript
await Word.run(async (context) => {
    const range = context.document.getSelection();
    const listFormat = range.listFormat;
    
    listFormat.load("isSingleList");
    await context.sync();
    
    if (listFormat.isSingleList) {
        console.log("The selection contains a single list.");
    } else {
        console.log("The selection contains multiple lists or no list.");
    }
});
```

---

### isSingleListTemplate

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Indicates whether the `ListFormat` object contains a single list template.

#### Examples

**Example**: Check if a paragraph's list formatting uses a single list template and display the result in the console.

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    const listFormat = paragraph.listFormat;
    
    listFormat.load("isSingleListTemplate");
    await context.sync();
    
    console.log("Uses single list template: " + listFormat.isSingleListTemplate);
});
```

---

### list

**Type:** `Word.List`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns a `List` object that represents the first formatted list contained in the `ListFormat` object.

#### Examples

**Example**: Get the list object from a selected range and display its ID in the console.

```typescript
await Word.run(async (context) => {
    const range = context.document.getSelection();
    const listFormat = range.listFormat;
    const list = listFormat.list;
    
    list.load("id");
    await context.sync();
    
    console.log("List ID: " + list.id);
});
```

---

### listLevelNumber

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the list level number for the first paragraph for the `ListFormat` object.

#### Examples

**Example**: Set the list level number to 2 for the selected paragraph to indent it as a second-level list item

```typescript
await Word.run(async (context) => {
    const range = context.document.getSelection();
    range.listFormat.listLevelNumber = 2;
    
    await context.sync();
});
```

---

### listString

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the string representation of the list value of the first paragraph in the range for the `ListFormat` object.

#### Examples

**Example**: Get and display the list string representation (e.g., "1.", "a.", "i.") of the first paragraph in the selected range.

```typescript
await Word.run(async (context) => {
    const range = context.document.getSelection();
    const listFormat = range.listFormat;
    
    listFormat.load("listString");
    await context.sync();
    
    console.log("List string: " + listFormat.listString);
});
```

---

### listTemplate

**Type:** `Word.ListTemplate`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the list template associated with the `ListFormat` object.

#### Examples

**Example**: Get and display the list template type of the first paragraph in the document to determine what kind of list formatting is applied.

```typescript
await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    const listFormat = firstParagraph.listOrNullObject;
    const listTemplate = listFormat.listTemplate;
    
    listFormat.load("listTemplate");
    await context.sync();
    
    if (listFormat.isNullObject) {
        console.log("The paragraph is not part of a list.");
    } else {
        console.log("List template type:", listTemplate);
    }
});
```

---

### listType

**Type:** `Word.ListType | "ListNoNumbering" | "ListListNumOnly" | "ListBullet" | "ListSimpleNumbering" | "ListOutlineNumbering" | "ListMixedNumbering" | "ListPictureBullet"`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the type of the list for the `ListFormat` object.

#### Examples

**Example**: Check if the selected text is formatted as a bulleted list and display the list type in the console.

```typescript
await Word.run(async (context) => {
    const range = context.document.getSelection();
    const listFormat = range.listFormat;
    
    // Load the listType property
    listFormat.load("listType");
    
    await context.sync();
    
    // Check and display the list type
    console.log("List type:", listFormat.listType);
    
    if (listFormat.listType === Word.ListType.bullet || 
        listFormat.listType === "ListBullet") {
        console.log("The selected text is a bulleted list");
    } else if (listFormat.listType === Word.ListType.noNumbering || 
               listFormat.listType === "ListNoNumbering") {
        console.log("The selected text is not in a list");
    } else {
        console.log("The selected text is in a numbered list");
    }
});
```

---

### listValue

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the numeric value of the the first paragraph in the range for the `ListFormat` object.

#### Examples

**Example**: Get and display the numeric value of the first paragraph in a numbered list

```typescript
await Word.run(async (context) => {
    const range = context.document.getSelection();
    const listFormat = range.listFormat;
    
    listFormat.load("listValue");
    await context.sync();
    
    console.log("List item number: " + listFormat.listValue);
});
```

---

## Methods

### applyBulletDefault

**Kind:** `configure`

Adds bullets and formatting to the paragraphs in the range.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `defaultListBehavior`: `Word.DefaultListBehavior` (optional)
    Optional. Specifies the default list behavior. Default is `DefaultListBehavior.word97`.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `defaultListBehavior`: `"Word97" | "Word2000" | "Word2002"` (optional)
    Optional. Specifies the default list behavior. Default is `DefaultListBehavior.word97`.

  **Returns:** `void`

#### Examples

**Example**: Apply default bullet formatting to all paragraphs in the first content control

```typescript
await Word.run(async (context) => {
    const contentControl = context.document.contentControls.getFirst();
    const range = contentControl.getRange();
    
    range.listFormat.applyBulletDefault(Word.DefaultListBehavior.respectCurrentList);
    
    await context.sync();
});
```

---

### applyListTemplateWithLevel

**Kind:** `configure`

Applies a list template with a specific level to the paragraphs in the range.

#### Signature

**Parameters:**
- `listTemplate`: `Word.ListTemplate` (required)
  The list template to apply.
- `options`: `Word.ListTemplateApplyOptions` (optional)
  Optional. Options for applying the list template, such as whether to continue the previous list or which part of the list to apply the template to.

**Returns:** `void`

#### Examples

**Example**: Apply a numbered list template at level 2 to all paragraphs in the first content control of the document.

```typescript
await Word.run(async (context) => {
    const contentControl = context.document.contentControls.getFirst();
    const range = contentControl.getRange();
    
    // Apply a numbered list template at level 2
    range.listFormat.applyListTemplateWithLevel(
        Word.ListTemplate.numberDefault,
        { level: 2 }
    );
    
    await context.sync();
});
```

---

### applyNumberDefault

**Kind:** `configure`

Adds numbering and formatting to the paragraphs in the range.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `defaultListBehavior`: `Word.DefaultListBehavior` (optional)
    Optional. Specifies the default list behavior.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `defaultListBehavior`: `"Word97" | "Word2000" | "Word2002"` (optional)
    Optional. Specifies the default list behavior.

  **Returns:** `void`

#### Examples

**Example**: Apply default numbered list formatting to all paragraphs in the first content control

```typescript
await Word.run(async (context) => {
    const contentControl = context.document.contentControls.getFirst();
    const range = contentControl.getRange();
    
    range.listFormat.applyNumberDefault(Word.DefaultListBehavior.respectCurrentList);
    
    await context.sync();
});
```

---

### applyOutlineNumberDefault

**Kind:** `configure`

Adds outline numbering and formatting to the paragraphs in the range.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `defaultListBehavior`: `Word.DefaultListBehavior` (optional)
    Optional. Specifies the default list behavior.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `defaultListBehavior`: `"Word97" | "Word2000" | "Word2002"` (optional)
    Optional. Specifies the default list behavior.

  **Returns:** `void`

#### Examples

**Example**: Apply default outline numbering to all paragraphs in the document body

```typescript
await Word.run(async (context) => {
    const body = context.document.body;
    const listFormat = body.listFormat;
    
    // Apply outline numbering with default list behavior
    listFormat.applyOutlineNumberDefault(Word.DefaultListBehavior.respectCurrentList);
    
    await context.sync();
});
```

---

### canContinuePreviousList

**Kind:** `read`

Determines whether the `ListFormat` object can continue a previous list.

#### Signature

**Parameters:**
- `listTemplate`: `Word.ListTemplate` (required)
  The list template to check.

**Returns:** `OfficeExtension.ClientResult<Word.Continue>`
A `Continue` value indicating whether continuation is possible.

#### Examples

**Example**: Check if the current paragraph can continue the formatting from a previous numbered list and display the result in a message.

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const paragraph = context.document.body.paragraphs.getFirst();
    const listFormat = paragraph.listFormat;
    
    // Get the list template from a previous list (e.g., the first list in the document)
    const lists = context.document.body.lists;
    lists.load("items");
    await context.sync();
    
    if (lists.items.length > 0) {
        const firstList = lists.items[0];
        const listTemplate = firstList.getLevelParagraphs(0).getFirst().listFormat;
        listTemplate.load("listTemplate");
        await context.sync();
        
        // Check if the current paragraph can continue the previous list
        const canContinue = listFormat.canContinuePreviousList(listTemplate.listTemplate);
        
        console.log(`Can continue previous list: ${canContinue}`);
    } else {
        console.log("No existing lists found in the document.");
    }
});
```

---

### convertNumbersToText

**Kind:** `configure`

Converts numbers in the list to plain text.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `numberType`: `Word.NumberType` (optional)
    Optional. The type of number to convert.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `numberType`: `"Paragraph" | "ListNum" | "AllNumbers"` (optional)
    Optional. The type of number to convert.

  **Returns:** `void`

#### Examples

**Example**: Convert all numbered list items in the first paragraph to plain text, removing the automatic numbering formatting

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const paragraph = context.document.body.paragraphs.getFirst();
    
    // Get the list format of the paragraph
    const listFormat = paragraph.listFormat;
    
    // Load the list level to verify it's a numbered list
    listFormat.load("listLevelNumber");
    
    await context.sync();
    
    // Convert the numbers to plain text
    listFormat.convertNumbersToText(Word.NumberType.arabic);
    
    await context.sync();
});
```

---

### countNumberedItems

**Kind:** `read`

Counts the numbered items in the list.

#### Signature

**Parameters:**
- `options`: `Word.ListFormatCountNumberedItemsOptions` (optional)
  Optional. Options for counting numbered items, such as the type of number and the level to count.

**Returns:** `OfficeExtension.ClientResult<number>`
The number of items.

#### Examples

**Example**: Count and display the total number of numbered items in the current document's selection

```typescript
await Word.run(async (context) => {
    const range = context.document.getSelection();
    const listFormat = range.listFormat;
    
    const count = listFormat.countNumberedItems();
    
    await context.sync();
    
    console.log(`Total numbered items: ${count.value}`);
});
```

---

### listIndent

**Kind:** `configure`

Indents the list by one level.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Indent the first paragraph in the document by one list level to create a nested list item

```typescript
await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    firstParagraph.listOrNullObject.listIndent();
    
    await context.sync();
});
```

---

### listOutdent

**Kind:** `configure`

Outdents the list by one level.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Decrease the indentation level of the first paragraph in the document by one level (outdent it in the list hierarchy)

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const firstParagraph = context.document.body.paragraphs.getFirst();
    
    // Get the list format of the paragraph
    const listFormat = firstParagraph.listOrNullObject;
    
    // Load the list level to check if it's part of a list
    listFormat.load("level");
    
    await context.sync();
    
    // Outdent the list item by one level
    if (!listFormat.isNullObject) {
        listFormat.listOutdent();
    }
    
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
  - `options`: `Word.Interfaces.ListFormatLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.ListFormat`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.ListFormat`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.ListFormat`

#### Examples

**Example**: Read and display the list level of the first paragraph in the document

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const paragraph = context.document.body.paragraphs.getFirst();
    
    // Get the list format of the paragraph
    const listFormat = paragraph.listFormat;
    
    // Load the list level property
    listFormat.load("listLevelNumber");
    
    // Sync to execute the load command
    await context.sync();
    
    // Display the list level (0-based, or undefined if not in a list)
    console.log("List level: " + listFormat.listLevelNumber);
});
```

---

### removeNumbers

**Kind:** `configure`

Removes numbering from the list.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `numberType`: `Word.NumberType` (optional)
    Optional. The type of number to remove.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `numberType`: `"Paragraph" | "ListNum" | "AllNumbers"` (optional)
    Optional. The type of number to remove.

  **Returns:** `void`

#### Examples

**Example**: Remove all numbering from the first numbered list in the document

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document that has list formatting
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("items");
    
    await context.sync();
    
    // Find and remove numbering from the first numbered list
    for (let i = 0; i < paragraphs.items.length; i++) {
        const listFormat = paragraphs.items[i].listOrNullObject;
        listFormat.load("levelNumber");
        
        await context.sync();
        
        if (listFormat.isNullObject === false) {
            // Remove the numbering from this list item
            listFormat.removeNumbers(Word.NumberType.arabic);
            break;
        }
    }
    
    await context.sync();
});
```

---

### set

**Kind:** `write`

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `Interfaces.ListFormatUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.ListFormat` (required)

  **Returns:** `void`

#### Examples

**Example**: Apply list formatting to a paragraph by setting multiple list properties at once, including list type and level indentation

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    const listFormat = paragraph.listOrNullObject;
    
    listFormat.set({
        level: 1,
        listString: "•"
    });
    
    await context.sync();
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.ListFormat` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.ListFormatData`) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.ListFormatData`

#### Examples

**Example**: Serialize a paragraph's list formatting properties to JSON for logging or data export purposes

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const paragraph = context.document.body.paragraphs.getFirst();
    const listFormat = paragraph.listFormat;
    
    // Load the list format properties
    listFormat.load("listLevelNumber,listString,listType");
    
    await context.sync();
    
    // Convert the list format to a plain JavaScript object
    const listFormatData = listFormat.toJSON();
    
    // Log or export the serialized data
    console.log("List Format Data:", JSON.stringify(listFormatData, null, 2));
    
    // Example output:
    // {
    //   "listLevelNumber": 0,
    //   "listString": "1.",
    //   "listType": "Number"
    // }
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.ListFormat`

#### Examples

**Example**: Track a list format object across multiple sync calls to maintain its reference while modifying list properties in different batches

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    const listFormat = paragraph.listFormat;
    
    // Track the list format object to use it across sync calls
    listFormat.track();
    
    await context.sync();
    
    // First batch: Set list type
    listFormat.setListType(Word.ListType.numbered);
    await context.sync();
    
    // Second batch: Modify list level (object reference still valid due to tracking)
    listFormat.level = 1;
    await context.sync();
    
    // Clean up: Untrack when done
    listFormat.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

#### Signature

**Returns:** `Word.ListFormat`

#### Examples

**Example**: Remove list formatting from a paragraph and untrack the ListFormat object to free memory after the operation is complete.

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const paragraph = context.document.body.paragraphs.getFirst();
    const listFormat = paragraph.listFormat;
    
    // Track the object to work with it
    listFormat.track();
    
    // Load properties to check list status
    listFormat.load("listLevelNumber");
    await context.sync();
    
    // Remove list formatting if present
    if (listFormat.listLevelNumber !== undefined) {
        listFormat.listLevelNumber = undefined;
    }
    
    // Untrack the object to release memory
    listFormat.untrack();
    
    await context.sync();
});
```

---

## Source

- /en-us/javascript/api/word
- /en-us/javascript/api/office/officeextension.clientobject
- /en-us/javascript/api/requirement-sets/word/word-api-requirement-sets
- /en-us/javascript/api/word/word.requestcontext
- /en-us/javascript/api/word/word.list
- /en-us/javascript/api/word/word.listtemplate
- /en-us/javascript/api/word/word.listtype
- /en-us/javascript/api/word/word.defaultlistbehavior
- /en-us/javascript/api/word/word.listtemplateapplyoptions
- /en-us/javascript/api/word/word.numbertype
- /en-us/javascript/api/office/officeextension.clientresult
- /en-us/javascript/api/word/word.continue
- /en-us/javascript/api/word/word.listformatcountnumbereditemsoptions
- /en-us/javascript/api/word/word.interfaces.listformatloadoptions
- /en-us/javascript/api/word/word.listformat
- /en-us/javascript/api/word/word.interfaces.listformatupdatedata
- /en-us/javascript/api/office/officeextension.updateoptions
- /en-us/javascript/api/word/word.interfaces.listformatdata
- /en-us/javascript/api/office/officeextension.clientrequestcontext
