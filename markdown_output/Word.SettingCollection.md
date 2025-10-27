# Word.SettingCollection

**Package:** `word`

**API Set:** WordApi 1.4

**Extends:** `OfficeExtension.ClientObject`

## Description

Contains the collection of [Word.Setting](/en-us/javascript/api/word/word.setting) objects.

## Class Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-settings.yaml

// Deletes all custom settings this add-in had set on this document.
await Word.run(async (context) => {
  const settings: Word.SettingCollection = context.document.settings;
  settings.deleteAll();
  await context.sync();
  console.log("All settings deleted.");
});
```

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a SettingCollection to verify the connection to the Word host application before performing operations on document settings.

```typescript
await Word.run(async (context) => {
    const settings = context.document.settings;
    
    // Access the request context associated with the SettingCollection
    const requestContext = settings.context;
    
    // Use the context to load properties and verify connection
    settings.load("items");
    await requestContext.sync();
    
    console.log("Connected to Word host application");
    console.log(`Number of settings: ${settings.items.length}`);
});
```

---

### items

**Type:** `Word.Setting[]`

Gets the loaded child items in this collection.

#### Examples

**Example**: Retrieve and display all custom settings that the add-in has stored in the current Word document.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-settings.yaml

// Gets all custom settings this add-in set on this document.
await Word.run(async (context) => {
  const settings: Word.SettingCollection = context.document.settings;
  settings.load("items");
  await context.sync();

  if (settings.items.length == 0) {
    console.log("There are no settings.");
  } else {
    console.log("All settings:");
    for (let i = 0; i < settings.items.length; i++) {
      console.log(settings.items[i]);
    }
  }
});
```

---

## Methods

### add

**Kind:** `create`

Creates a new setting or sets an existing setting.

#### Signature

**Parameters:**
- `key`: `string` (required)
  The setting's key, which is case-sensitive.
- `value`: `any` (required)
  The setting's value.

**Returns:** `Word.Setting`

#### Examples

**Example**: Add a new custom setting to the Word document with a specified key-value pair, or update the value if the key already exists.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-settings.yaml

// Adds a new custom setting or
// edits the value of an existing one.
await Word.run(async (context) => {
  const key = (document.getElementById("key") as HTMLInputElement).value;
  if (key == "") {
    console.error("Key shouldn't be empty.");
    return;
  }

  const value = (document.getElementById("value") as HTMLInputElement).value;
  const settings: Word.SettingCollection = context.document.settings;
  const setting: Word.Setting = settings.add(key, value);
  setting.load();
  await context.sync();

  console.log("Setting added or edited:", setting);
});
```

---

### deleteAll

**Kind:** `delete`

Deletes all settings in this add-in.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Demonstrate adding a custom setting to the document, verifying the setting count, deleting all settings, and confirming the settings collection is empty.

```typescript
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Queue commands add a setting.
    const settings = context.document.settings;
    settings.add('startMonth', { month: 'March', year: 1998 });

    // Queue a command to get the count of settings.
    const count = settings.getCount();

    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    await context.sync();
    console.log(count.value);

    // Queue a command to delete all settings.
    settings.deleteAll();

    // Queue a command to get the new count of settings.
    count = settings.getCount();

    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    await context.sync();
    console.log(count.value);
});
```

---

### getCount

**Kind:** `read`

Gets the count of settings.

#### Signature

**Returns:** `OfficeExtension.ClientResult<number>`

#### Examples

**Example**: Demonstrate retrieving the count of document settings before and after adding a setting and then deleting all settings.

```typescript
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Queue commands add a setting.
    const settings = context.document.settings;
    settings.add('startMonth', { month: 'March', year: 1998 });

    // Queue a command to get the count of settings.
    const count = settings.getCount();

    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    await context.sync();
    console.log(count.value);

    // Queue a command to delete all settings.
    settings.deleteAll();

    // Queue a command to get the new count of settings.
    count = settings.getCount();

    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    await context.sync();
    console.log(count.value);
});
```

---

### getItem

**Kind:** `read`

Gets a setting object by its key, which is case-sensitive. Throws an `ItemNotFound` error if the setting doesn't exist.

#### Signature

**Parameters:**
- `key`: `string` (required)
  The key that identifies the setting object.

**Returns:** `Word.Setting`

#### Examples

**Example**: Add a custom setting named 'startMonth' to the document, retrieve it, and log its value to the console.

```typescript
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Queue commands add a setting.
    const settings = context.document.settings;
    settings.add('startMonth', { month: 'March', year: 1998 });

    // Queue a command to retrieve a setting.
    const startMonth = settings.getItem('startMonth');

    // Queue a command to load the setting.
    startMonth.load();

    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    await context.sync();
    console.log(JSON.stringify(startMonth.value));
});
```

---

### getItemOrNullObject

**Kind:** `read`

Gets a setting object by its key, which is case-sensitive. If the setting doesn't exist, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

#### Signature

**Parameters:**
- `key`: `string` (required)
  The key that identifies the setting object.

**Returns:** `Word.Setting`

#### Examples

**Example**: Retrieve and display the values of 'startMonth' and 'endMonth' settings from the document, handling cases where either setting may not exist.

```typescript
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Queue commands add a setting.
    const settings = context.document.settings;
    settings.add('startMonth', { month: 'March', year: 1998 });
    
    // Queue commands to retrieve settings.
    const startMonth = settings.getItemOrNullObject('startMonth');
    const endMonth = settings.getItemOrNullObject('endMonth');

    // Queue commands to load settings.
    startMonth.load();
    endMonth.load();

    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    await context.sync();
        if (startMonth.isNullObject) {
            console.log("No such setting.");
        }
        else {
            console.log(JSON.stringify(startMonth.value));
        }
        if (endMonth.isNullObject) {
            console.log("No such setting.");
        }
        else {
            console.log(JSON.stringify(endMonth.value));
        }
});
```

---

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.SettingCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.SettingCollection`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.SettingCollection`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `OfficeExtension.LoadOption` (optional)
    `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.SettingCollection`

#### Examples

**Example**: Load and display all custom document settings with their keys and values

```typescript
await Word.run(async (context) => {
    // Get the settings collection
    const settings = context.document.settings;
    
    // Load the settings collection with key and value properties
    settings.load("items");
    
    await context.sync();
    
    // Display all settings
    console.log(`Total settings: ${settings.items.length}`);
    settings.items.forEach(setting => {
        console.log(`Key: ${setting.key}, Value: ${setting.value}`);
    });
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.SettingCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.SettingCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

#### Signature

**Returns:** `Word.Interfaces.SettingCollectionData`

#### Examples

**Example**: Serialize a collection of document settings to JSON format for logging or external storage

```typescript
await Word.run(async (context) => {
    // Get the settings collection
    const settings = context.document.settings;
    
    // Load the settings collection with their properties
    settings.load("key,value");
    
    await context.sync();
    
    // Convert the settings collection to a plain JavaScript object
    const settingsJSON = settings.toJSON();
    
    // Log the JSON representation (can be used for debugging or storage)
    console.log("Settings as JSON:", JSON.stringify(settingsJSON, null, 2));
    
    // Access the items array from the JSON object
    console.log("Number of settings:", settingsJSON.items.length);
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.SettingCollection`

#### Examples

**Example**: Track custom document settings across multiple sync calls to monitor and update a document's metadata without encountering InvalidObjectPath errors

```typescript
await Word.run(async (context) => {
    // Get the settings collection
    const settings = context.document.settings;
    
    // Track the settings collection to use it across multiple sync calls
    settings.track();
    
    // Load the settings
    settings.load("items");
    await context.sync();
    
    // First sync - read existing settings
    console.log(`Found ${settings.items.length} settings`);
    
    // Add a new setting
    settings.add("lastModifiedBy", "John Doe");
    await context.sync();
    
    // Second sync - the tracked object remains valid
    // Without track(), accessing settings here might cause InvalidObjectPath error
    settings.add("documentVersion", "1.0");
    await context.sync();
    
    // Third sync - still valid because it's tracked
    settings.load("items");
    await context.sync();
    console.log(`Now have ${settings.items.length} settings`);
    
    // Untrack when done to free up resources
    settings.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

#### Signature

**Returns:** `Word.SettingCollection`

#### Examples

**Example**: Load and process document settings, then untrack the settings collection to free memory after use

```typescript
await Word.run(async (context) => {
    // Load the settings collection
    const settings = context.document.settings;
    settings.load("items");
    
    await context.sync();
    
    // Process the settings (e.g., read values)
    console.log(`Found ${settings.items.length} settings`);
    settings.items.forEach(setting => {
        console.log(`${setting.key}: ${setting.value}`);
    });
    
    // Untrack the collection to release memory
    settings.untrack();
    
    await context.sync();
});
```

---

## Source

- https://docs.microsoft.com/en-us/javascript/api/word
