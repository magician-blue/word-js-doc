# Word.ListTemplate

**Package:** `word`

**API Set:** WordApiDesktop 1.1

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents a list template.

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

### context

**Type:** `None`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a ListTemplate object to verify the connection between the add-in and Word application

```typescript
await Word.run(async (context) => {
    // Get the first list in the document
    const lists = context.document.body.lists;
    lists.load("items");
    await context.sync();
    
    if (lists.items.length > 0) {
        // Get the list template from the first list
        const listTemplate = lists.items[0].getListTemplate();
        
        // Access the request context associated with the list template
        const requestContext = listTemplate.context;
        
        // Use the context to perform operations
        listTemplate.load("levelTypes");
        await requestContext.sync();
        
        console.log("List template context is connected to Word application");
        console.log("Level types:", listTemplate.levelTypes);
    }
});
```

---

### listLevels

**Type:** `None`

Gets a ListLevelCollection object that represents all the levels for the list template.

#### Examples

**Example**: Get all list levels from a list template and log the count of available levels to the console.

```typescript
await Word.run(async (context) => {
    // Get the first list in the document
    const lists = context.document.body.lists;
    lists.load("items");
    await context.sync();
    
    if (lists.items.length > 0) {
        const firstList = lists.items[0];
        
        // Get the list template
        const listTemplate = firstList.listTemplate;
        
        // Access the listLevels property
        const listLevels = listTemplate.listLevels;
        listLevels.load("items");
        await context.sync();
        
        console.log(`Number of list levels: ${listLevels.items.length}`);
        
        // Optionally iterate through each level
        listLevels.items.forEach((level, index) => {
            console.log(`Level ${index + 1} exists`);
        });
    }
});
```

---

### outlineNumbered

**Type:** `None`

Specifies whether the list template is outline numbered.

#### Examples

**Example**: Check if a list template is outline numbered and display the result in the console.

```typescript
await Word.run(async (context) => {
    // Get the first list in the document
    const lists = context.document.body.lists;
    lists.load("items");
    await context.sync();
    
    if (lists.items.length > 0) {
        const firstList = lists.items[0];
        const listTemplate = firstList.getListTemplate();
        listTemplate.load("outlineNumbered");
        await context.sync();
        
        console.log(`Is outline numbered: ${listTemplate.outlineNumbered}`);
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

**Example**: Load and display the level count property of the first list template in the document

```typescript
await Word.run(async (context) => {
    // Get the first list in the document
    const lists = context.document.body.lists;
    lists.load("items");
    await context.sync();
    
    if (lists.items.length > 0) {
        // Get the list template from the first list
        const listTemplate = lists.items[0].getListTemplate();
        
        // Load the levelCount property of the list template
        listTemplate.load("levelCount");
        await context.sync();
        
        // Display the loaded property
        console.log(`List template has ${listTemplate.levelCount} levels`);
    } else {
        console.log("No lists found in the document");
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
  - `properties`: `None` (required)
  - `options`: `None` (required)

  **Returns:** `None`

**Overload 2:**

  **Parameters:**
  - `properties`: `None` (required)

  **Returns:** `None`

#### Examples

**Example**: Configure multiple properties of a list template at once, including its bullet character and font settings for the first level.

```typescript
await Word.run(async (context) => {
    // Get the first list in the document
    const lists = context.document.body.lists;
    lists.load("items");
    await context.sync();
    
    if (lists.items.length > 0) {
        const list = lists.items[0];
        const listTemplate = list.listTemplate;
        
        // Set multiple properties of the list template at once
        listTemplate.set({
            bulletFormat: "●",
            fontName: "Arial",
            fontSize: 12
        });
        
        await context.sync();
        console.log("List template properties updated successfully");
    }
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.ListTemplate object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.ListTemplateData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Serialize a list template to JSON format for logging or data transfer purposes

```typescript
await Word.run(async (context) => {
    // Get the first list in the document
    const lists = context.document.body.lists;
    lists.load("items");
    await context.sync();
    
    if (lists.items.length > 0) {
        // Get the list template from the first list
        const listTemplate = lists.items[0].getListTemplate();
        listTemplate.load("*");
        await context.sync();
        
        // Convert the list template to a plain JSON object
        const listTemplateJSON = listTemplate.toJSON();
        
        // Log the JSON representation
        console.log("List Template JSON:", JSON.stringify(listTemplateJSON, null, 2));
    }
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to

#### Signature

**Returns:** `None`

#### Examples

**Example**: Track a list template object to maintain its reference across multiple sync calls when modifying list properties in different batches

```typescript
await Word.run(async (context) => {
    // Get the first list in the document
    const list = context.document.body.lists.getFirst();
    const listTemplate = list.getListTemplate();
    
    // Track the list template to use it across sync calls
    listTemplate.track();
    
    // Load properties
    listTemplate.load("id");
    await context.sync();
    
    console.log("List template ID: " + listTemplate.id);
    
    // Perform another sync operation - the tracked object remains valid
    await context.sync();
    
    // Can still access the list template properties after multiple syncs
    console.log("Still accessible: " + listTemplate.id);
    
    // Untrack when done to free up memory
    listTemplate.untrack();
    
    await context.sync();
});
```

---

## Source

- https://docs.microsoft.com/en-us/javascript/api/word/word.listtemplate
