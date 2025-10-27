# Word.TemplateCollection

**Package:** `word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `OfficeExtension.ClientObject`

## Description

Contains a collection of [Word.Template](/en-us/javascript/api/word/word.template) objects that represent all the templates that are currently available. This collection includes open templates, templates attached to open documents, and global templates loaded in the Templates and Add-ins dialog box. To learn how to access this dialog in the Word UI, see Load or unload a template or add-in program: https://support.microsoft.com/office/2479fe53-f849-4394-88bb-2a6e2a39479d.

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from the TemplateCollection to verify the connection between the add-in and Word application before performing template operations.

```typescript
await Word.run(async (context) => {
    const templates = context.application.templates;
    
    // Access the request context associated with the TemplateCollection
    const requestContext = templates.context;
    
    // Verify the context is valid by checking if it matches the current context
    if (requestContext === context) {
        console.log("TemplateCollection is properly connected to the Word application context");
    }
    
    // Use the context to load and sync template data
    templates.load("items");
    await context.sync();
    
    console.log(`Number of templates available: ${templates.items.length}`);
});
```

---

### items

**Type:** `Word.Template[]`

Gets the loaded child items in this collection.

#### Examples

**Example**: Get all available templates and display their names in the console.

```typescript
await Word.run(async (context) => {
    const templateCollection = context.application.templates;
    templateCollection.load("items");
    await context.sync();
    
    const templates = templateCollection.items;
    console.log(`Found ${templates.length} template(s):`);
    
    for (let i = 0; i < templates.length; i++) {
        templates[i].load("name");
    }
    await context.sync();
    
    templates.forEach((template, index) => {
        console.log(`${index + 1}. ${template.name}`);
    });
});
```

---

## Methods

### getCount

**Kind:** `read`

Returns the number of items in the collection.

#### Signature

**Returns:** `OfficeExtension.ClientResult<number>`

#### Examples

**Example**: Get the total number of templates currently available in Word and display it in the console

```typescript
await Word.run(async (context) => {
    const templates = context.application.templates;
    const templateCount = templates.getCount();
    
    await context.sync();
    
    console.log(`Total number of templates available: ${templateCount.value}`);
});
```

---

### getItemAt

**Kind:** `read`

Gets a Template object by its index in the collection.

#### Signature

**Parameters:**
- `index`: `number` (required)
  The index of the template to retrieve.

**Returns:** `Word.Template`

#### Examples

**Example**: Get the first template from the templates collection and display its name in the console.

```typescript
await Word.run(async (context) => {
    // Get the templates collection
    const templates = context.application.templates;
    
    // Get the first template (at index 0)
    const firstTemplate = templates.getItemAt(0);
    
    // Load the name property
    firstTemplate.load("name");
    
    // Sync to execute the queued commands
    await context.sync();
    
    // Display the template name
    console.log("First template name: " + firstTemplate.name);
});
```

---

### importBuildingBlocks

**Kind:** `load`

Imports the building blocks for all templates into Microsoft Word.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Import all building blocks from available templates into Word to make them accessible for use in the current document.

```typescript
await Word.run(async (context) => {
    // Get the template collection
    const templates = context.application.templates;
    
    // Import building blocks from all templates
    templates.importBuildingBlocks();
    
    // Sync to execute the import operation
    await context.sync();
    
    console.log("Building blocks imported successfully from all templates.");
});
```

---

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.TemplateCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.TemplateCollection`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.TemplateCollection`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `OfficeExtension.LoadOption` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.TemplateCollection`

#### Examples

**Example**: Load and display the names of all available templates in the Word application

```typescript
await Word.run(async (context) => {
    // Get the template collection
    const templates = context.application.templates;
    
    // Load the name property for all templates in the collection
    templates.load("items/name");
    
    // Synchronize the document state
    await context.sync();
    
    // Display the template names
    console.log("Available templates:");
    for (let i = 0; i < templates.items.length; i++) {
        console.log(`- ${templates.items[i].name}`);
    }
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.TemplateCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.TemplateCollectionData) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

#### Signature

**Returns:** `Word.Interfaces.TemplateCollectionData`

#### Examples

**Example**: Serialize the template collection to JSON format for logging or debugging purposes to inspect available templates and their properties.

```typescript
await Word.run(async (context) => {
    // Get the template collection
    const templates = context.application.templates;
    
    // Load properties we want to inspect
    templates.load("items");
    
    await context.sync();
    
    // Convert the collection to a plain JavaScript object
    const templatesJSON = templates.toJSON();
    
    // Log the JSON representation
    console.log("Templates available:");
    console.log(JSON.stringify(templatesJSON, null, 2));
    
    // Access the items array from the JSON object
    console.log(`Total templates: ${templatesJSON.items.length}`);
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.TemplateCollection`

#### Examples

**Example**: Track a template collection across multiple sync calls to safely access template properties without getting "InvalidObjectPath" errors

```typescript
await Word.run(async (context) => {
    const templates = context.application.templates;
    templates.load("items");
    await context.sync();
    
    // Track the collection to use it across multiple sync calls
    templates.track();
    
    // First sync - get initial data
    console.log(`Found ${templates.items.length} templates`);
    await context.sync();
    
    // Second sync - can still safely access the tracked collection
    for (let i = 0; i < templates.items.length; i++) {
        const template = templates.items[i];
        template.load("name");
    }
    await context.sync();
    
    // Third sync - still valid because collection is tracked
    templates.items.forEach(template => {
        console.log(`Template: ${template.name}`);
    });
    
    // Clean up tracking when done
    templates.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.TemplateCollection`

#### Examples

**Example**: Load the template collection, access its items, and then release the memory associated with the collection after use to optimize performance.

```typescript
await Word.run(async (context) => {
    // Load the template collection
    const templates = context.application.templates;
    templates.load("items");
    
    await context.sync();
    
    // Use the template collection (e.g., log count)
    console.log(`Number of templates: ${templates.items.length}`);
    
    // Release memory associated with the template collection
    templates.untrack();
    
    await context.sync();
});
```

---

## Source

- /en-us/javascript/api/word
- /en-us/javascript/api/office/officeextension.clientobject
- /en-us/javascript/api/requirement-sets/word/word-api-requirement-sets
- /en-us/javascript/api/word/word.template
- https://support.microsoft.com/office/2479fe53-f849-4394-88bb-2a6e2a39479d
- /en-us/javascript/api/word/word.requestcontext
- /en-us/javascript/api/office/officeextension.clientresult
- /en-us/javascript/api/word/word.interfaces.templatecollectionloadoptions
- /en-us/javascript/api/word/word.interfaces.collectionloadoptions
- /en-us/javascript/api/word/word.templatecollection
- /en-us/javascript/api/office/officeextension.loadoption
- /en-us/javascript/api/word/word.interfaces.templatecollectiondata
- /en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member
