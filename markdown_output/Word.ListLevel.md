# Word.ListLevel

**Package:** `https://learn.microsoft.com/en-us/javascript/api/word`

**API Set:** WordApiDesktop 1.1

**Extends:** `officeextension.clientobject`

## Description

Represents a list level.

## Class Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/20-lists/manage-list-styles.yaml

// Gets the properties of the specified style.
await Word.run(async (context) => {
  const styleName = (document.getElementById("style-name-to-use") as HTMLInputElement).value;
  if (styleName == "") {
    console.warn("Enter a style name to get properties.");
    return;
  }

  const style: Word.Style = context.document.getStyles().getByNameOrNullObject(styleName);
  style.load("type");
  await context.sync();

  if (style.isNullObject || style.type != Word.StyleType.list) {
    console.warn(`There's no existing style with the name '${styleName}'. Or this isn't a list style.`);
  } else {
    // Load objects to log properties and their values in the console.
    style.load();
    style.listTemplate.load();
    await context.sync();

    console.log(`Properties of the '${styleName}' style:`, style);

    const listLevels = style.listTemplate.listLevels;
    listLevels.load("items");
    await context.sync();

    console.log(`List levels of the '${styleName}' style:`, listLevels);
  }
});
```

## Properties

### alignment

**Type:** `None`

Specifies the horizontal alignment of the list level. The value can be 'Left', 'Centered', or 'Right'.

#### Examples

**Example**: Set the alignment of the first list level to centered for the first list in the document

```typescript
await Word.run(async (context) => {
    // Get the first list in the document
    const list = context.document.body.lists.getFirst();
    
    // Get the first list level
    const listLevel = list.levelTypes[0];
    
    // Set the alignment to centered
    listLevel.alignment = Word.Alignment.centered;
    
    await context.sync();
});
```

---

### context

**Type:** `None`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a ListLevel object to verify the connection between the add-in and Word application before performing operations on the list level.

```typescript
await Word.run(async (context) => {
    // Get the first list in the document
    const lists = context.document.body.lists;
    lists.load("items");
    await context.sync();
    
    if (lists.items.length > 0) {
        // Get the first level of the first list
        const listLevel = lists.items[0].getLevelParagraphs(1);
        listLevel.load("items");
        await context.sync();
        
        if (listLevel.items.length > 0) {
            const level = listLevel.items[0].listItem.level;
            
            // Access the context property to verify the connection
            const requestContext = level.context;
            
            // Use the context to perform operations
            console.log("Request context is connected:", requestContext !== null);
        }
    }
});
```

---

### font

**Type:** `None`

Gets a Font object that represents the character formatting of the specified object.

#### Examples

**Example**: Set the list level font to bold, blue, and 14pt size for the first list in the document.

```typescript
await Word.run(async (context) => {
    const firstList = context.document.body.lists.getFirst();
    const listLevel = firstList.getLevelParagraphs(0).getFirst().listItem.level;
    
    const font = listLevel.font;
    font.bold = true;
    font.color = "blue";
    font.size = 14;
    
    await context.sync();
});
```

---

### linkedStyle

**Type:** `None`

Specifies the name of the style that's linked to the specified list level object.

#### Examples

**Example**: Link a list level to a character style named "ListBullet" to apply consistent formatting across list items.

```typescript
await Word.run(async (context) => {
    const list = context.document.body.lists.getFirst();
    const listLevel = list.getLevelParagraphs(0).getFirst().listItem.level;
    
    // Link the list level to a style named "ListBullet"
    listLevel.linkedStyle = "ListBullet";
    
    await context.sync();
    console.log("List level linked to style: ListBullet");
});
```

---

### numberFormat

**Type:** `None`

Specifies the number format for the specified list level.

#### Examples

**Example**: Set the number format of the first list level to use uppercase Roman numerals (I, II, III, etc.)

```typescript
await Word.run(async (context) => {
    const list = context.document.body.lists.getFirst();
    const listLevel = list.getLevelParagraphs(0).getFirst().listItem.level;
    
    listLevel.numberFormat = "%I.";
    
    await context.sync();
});
```

---

### numberPosition

**Type:** `None`

Specifies the position (in points) of the number or bullet for the specified list level object.

#### Examples

**Example**: Set the number position to 36 points for the first level of the first list in the document

```typescript
await Word.run(async (context) => {
    const lists = context.document.body.lists;
    lists.load("items");
    await context.sync();
    
    if (lists.items.length > 0) {
        const firstList = lists.items[0];
        const listLevel = firstList.getLevelParagraphs(0).getFirst().listItem.level;
        listLevel.numberPosition = 36;
        
        await context.sync();
    }
});
```

---

### numberStyle

**Type:** `None`

Specifies the number style for the list level object.

#### Examples

**Example**: Set the number style of the first list level to uppercase Roman numerals (I, II, III, IV, etc.)

```typescript
await Word.run(async (context) => {
    const list = context.document.body.lists.getFirst();
    const listLevel = list.getLevelParagraphs(0).getFirst().listItem.level;
    
    listLevel.numberStyle = Word.ListNumberStyle.uppercaseRoman;
    
    await context.sync();
});
```

---

### resetOnHigher

**Type:** `None`

Specifies the list level that must appear before the specified list level restarts numbering at 1.

#### Examples

**Example**: Configure a list level so that it restarts numbering at 1 whenever a higher-level list item (level 0) appears before it.

```typescript
await Word.run(async (context) => {
    const list = context.document.body.lists.getFirst();
    const listLevel = list.getLevelParagraphs(1).getFirst().listItem.level;
    
    // Set level 1 to restart numbering when level 0 appears
    listLevel.resetOnHigher = 0;
    
    await context.sync();
    console.log("List level 1 will now restart numbering when level 0 appears");
});
```

---

### startAt

**Type:** `None`

Specifies the starting number for the specified list level object.

#### Examples

**Example**: Set the starting number of the first list level to 5 so the numbered list begins at 5 instead of 1

```typescript
await Word.run(async (context) => {
    const list = context.document.body.lists.getFirst();
    const listLevel = list.getLevelParagraphs(0).getFirst().listItem.level;
    
    listLevel.startAt = 5;
    
    await context.sync();
});
```

---

### tabPosition

**Type:** `None`

Specifies the tab position for the specified list level object.

#### Examples

**Example**: Set the tab position to 1 inch (72 points) for the first level of the first list in the document

```typescript
await Word.run(async (context) => {
    const lists = context.document.body.lists;
    lists.load("items");
    await context.sync();
    
    if (lists.items.length > 0) {
        const firstList = lists.items[0];
        const listLevel = firstList.getLevelParagraphs(0).getFirst().listItem.level;
        listLevel.tabPosition = 72; // 72 points = 1 inch
        
        await context.sync();
    }
});
```

---

### textPosition

**Type:** `None`

Specifies the position (in points) for the second line of wrapping text for the specified list level object.

#### Examples

**Example**: Set the text position (second line indent) to 36 points for the first level of the first list in the document

```typescript
await Word.run(async (context) => {
    const lists = context.document.body.lists;
    lists.load("items");
    await context.sync();
    
    if (lists.items.length > 0) {
        const firstList = lists.items[0];
        const listLevel = firstList.getLevelParagraphs(0).getFirst().listItem.level;
        listLevel.textPosition = 36;
        
        await context.sync();
    }
});
```

---

### trailingCharacter

**Type:** `None`

Specifies the character inserted after the number for the specified list level.

#### Examples

**Example**: Set the trailing character after the number to a tab for the first level of the first list in the document

```typescript
await Word.run(async (context) => {
    // Get the first list in the document
    const lists = context.document.body.lists;
    lists.load("items");
    await context.sync();
    
    if (lists.items.length > 0) {
        const firstList = lists.items[0];
        
        // Get the first level (level 1) of the list
        const listLevel = firstList.getLevelParagraphs(1);
        listLevel.load("items");
        await context.sync();
        
        if (listLevel.items.length > 0) {
            // Set the trailing character to tab
            listLevel.items[0].listItem.level = 1;
            const level1 = firstList.getLevelString(1);
            
            // Access the list level and set trailing character
            const paragraph = listLevel.items[0];
            paragraph.listItem.listLevel.trailingCharacter = Word.ListTrailingCharacter.tab;
            
            await context.sync();
        }
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

**Example**: Load and display the alignment property of the first list level in the first list of the document

```typescript
await Word.run(async (context) => {
    // Get the first list in the document
    const lists = context.document.body.lists;
    lists.load("items");
    await context.sync();
    
    if (lists.items.length > 0) {
        const firstList = lists.items[0];
        
        // Get the first list level
        const listLevel = firstList.getLevelParagraphs(0).getFirst().listItem.level;
        
        // Load the alignment property
        listLevel.load("alignment");
        await context.sync();
        
        // Display the alignment
        console.log("List level alignment: " + listLevel.alignment);
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

**Example**: Configure multiple formatting properties of the first list level in the first list, setting its alignment to right, number format to uppercase letters, and starting number to 5.

```typescript
await Word.run(async (context) => {
    // Get the first list in the document
    const list = context.document.body.lists.getFirst();
    
    // Get the first level of the list
    const listLevel = list.getLevelParagraphs(0).getFirst().listItem.level;
    
    // Set multiple properties at once
    listLevel.set({
        alignment: Word.Alignment.right,
        numberFormat: "%1",
        startAt: 5
    });
    
    await context.sync();
    console.log("List level properties updated successfully");
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.ListLevel object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.ListLevelData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Get the list level properties as a plain JavaScript object and log it to the console for debugging purposes.

```typescript
await Word.run(async (context) => {
    // Get the first list in the document
    const lists = context.document.body.lists;
    lists.load("items");
    await context.sync();
    
    if (lists.items.length > 0) {
        // Get the first level of the first list
        const listLevel = lists.items[0].getLevelParagraphs(0).getFirst().listItem.level;
        listLevel.load("alignment,font/name,font/size,numberFormat,numberPosition,startAt");
        await context.sync();
        
        // Convert to plain JavaScript object
        const listLevelData = listLevel.toJSON();
        
        // Log the plain object (useful for debugging or data export)
        console.log("List Level Data:", listLevelData);
        console.log("Number Format:", listLevelData.numberFormat);
        console.log("Start At:", listLevelData.startAt);
    }
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the par

#### Signature

**Returns:** `None`

#### Examples

**Example**: Track a list level object to maintain its reference across multiple sync calls when modifying list formatting properties in different batches

```typescript
await Word.run(async (context) => {
    const list = context.document.body.lists.getFirst();
    const listLevel = list.getLevelParagraphs(0).getFirst().listItem.level;
    
    // Track the list level object to use it across sync calls
    listLevel.track();
    
    await context.sync();
    
    // First batch: modify alignment
    listLevel.alignment = Word.Alignment.left;
    await context.sync();
    
    // Second batch: modify number format (object remains valid because it's tracked)
    listLevel.numberFormat = "1)";
    await context.sync();
    
    // Untrack when done to free up memory
    listLevel.untrack();
});
```

---

## Source

- https://learn.microsoft.com/en-us/javascript/api/word/word.listlevel
