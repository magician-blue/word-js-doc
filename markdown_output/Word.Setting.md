# Setting

**Package:** `word`

**API Set:** WordApi 1.4 None

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents a setting of the add-in.

## Class Examples

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

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a Setting object to verify the connection between the add-in and Word before performing operations

```typescript
await Word.run(async (context) => {
    // Get a setting object
    const settings = context.document.settings;
    const setting = settings.getItemOrNullObject("myCustomSetting");
    
    // Load the setting
    setting.load("key,value");
    await context.sync();
    
    // Access the request context from the setting object
    const settingContext = setting.context;
    
    // Verify the context is valid and connected
    if (settingContext && settingContext.document) {
        console.log("Setting's context is connected to Word document");
        
        // Use the context to perform additional operations
        await settingContext.sync();
    }
});
```

---

### key

**Type:** `string`

**Since:** WordApi 1.4

Gets the key of the setting.

#### Examples

**Example**: Add or update a custom document setting with a specified key-value pair obtained from user input fields.

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

### value

**Type:** `any`

**Since:** WordApi 1.4

Specifies the value of the setting.

#### Examples

**Example**: Add a new custom setting with a specified key-value pair to the Word document, or update the value if the key already exists.

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

## Methods

### delete

**Kind:** `delete`

Deletes the setting.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Delete a custom setting from the Word document and verify the setting count decreases after deletion.

```typescript
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Queue commands add a setting.
    const settings = context.document.settings;
    const startMonth = settings.add('startMonth', { month: 'March', year: 1998 });

    // Queue a command to get the count of settings.
    const count = settings.getCount();

    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    await context.sync();
    console.log(count.value);

    // Queue a command to delete the setting.
    startMonth.delete();

    // Queue a command to get the new count of settings.
    count = settings.getCount();

    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    await context.sync();
    console.log(count.value);
});
```

---

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.SettingLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.Setting`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.Setting`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.Setting`

#### Examples

**Example**: Load and read the value of an add-in setting named "userPreference"

```typescript
await Word.run(async (context) => {
    const settings = context.document.settings;
    const setting = settings.getItemOrNullObject("userPreference");
    
    setting.load("value");
    
    await context.sync();
    
    if (!setting.isNullObject) {
        console.log("Setting value:", setting.value);
    } else {
        console.log("Setting not found");
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
  - `properties`: `Interfaces.SettingUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.Setting` (required)

  **Returns:** `void`

#### Examples

**Example**: Configure multiple properties of an add-in setting at once, including its key, value, and metadata

```typescript
await Word.run(async (context) => {
    const settings = context.document.settings;
    const setting = settings.add("userPreferences", "");
    
    // Set multiple properties at once
    setting.set({
        value: JSON.stringify({ theme: "dark", fontSize: 14 }),
        key: "userPreferences"
    });
    
    await context.sync();
    console.log("Setting properties configured successfully");
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.Setting` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.SettingData`) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.SettingData`

#### Examples

**Example**: Retrieve a custom setting value and serialize it to JSON format for logging or external storage

```typescript
await Word.run(async (context) => {
    // Get a custom setting by key
    const setting = context.document.settings.getItemOrNullObject("myCustomSetting");
    setting.load("key,value");
    
    await context.sync();
    
    if (!setting.isNullObject) {
        // Convert the Setting object to a plain JavaScript object
        const settingData = setting.toJSON();
        
        // Now you can easily serialize it or log it
        console.log(JSON.stringify(settingData));
        // Output example: {"key":"myCustomSetting","value":"someValue"}
        
        // The plain object can be easily stored or transmitted
        const jsonString = JSON.stringify(settingData, null, 2);
        console.log("Setting as JSON:", jsonString);
    }
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.Setting`

#### Examples

**Example**: Track a custom XML part setting object across multiple sync calls to monitor and update its value without encountering InvalidObjectPath errors

```typescript
await Word.run(async (context) => {
    // Load a setting from the document
    const settings = context.document.settings;
    const setting = settings.getItemOrNullObject("MyCustomSetting");
    
    // Track the setting object for use across multiple sync calls
    setting.track();
    
    await context.sync();
    
    // First sync - check if setting exists
    if (!setting.isNullObject) {
        console.log("Current value: " + setting.value);
    } else {
        // Create the setting if it doesn't exist
        settings.add("MyCustomSetting", "initial value");
    }
    
    await context.sync();
    
    // Second sync - update the setting value
    // Without track(), this would throw InvalidObjectPath error
    setting.value = "updated value";
    
    await context.sync();
    
    console.log("Setting updated successfully");
    
    // Untrack when done to free up memory
    setting.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

#### Signature

**Returns:** `Word.Setting`

#### Examples

**Example**: Load a setting value, use it, then untrack it to free memory when done

```typescript
await Word.run(async (context) => {
    // Load a setting
    const setting = context.document.settings.getItemOrNullObject("myCustomSetting");
    setting.track();
    setting.load("value");
    
    await context.sync();
    
    if (!setting.isNullObject) {
        // Use the setting value
        console.log("Setting value:", setting.value);
        
        // Untrack the setting to release memory
        setting.untrack();
        
        await context.sync();
    }
});
```

---

## Source

- https://docs.microsoft.com/en-us/javascript/api/word
- https://docs.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets
- https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-settings.yaml
