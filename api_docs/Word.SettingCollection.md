# Word.SettingCollection class

Package: [word](/en-us/javascript/api/word)

Contains the collection of [Word.Setting](/en-us/javascript/api/word/word.setting) objects.

Extends: [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks
[ API set: WordApi 1.4 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples
```TypeScript
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
- [context](#context)  
  The request context associated with the object. This connects the add-in's process to the Office host application's process.
- [items](#items)  
  Gets the loaded child items in this collection.

## Methods
- [add(key, value)](#addkey-value)  
  Creates a new setting or sets an existing setting.
- [deleteAll()](#deleteall)  
  Deletes all settings in this add-in.
- [getCount()](#getcount)  
  Gets the count of settings.
- [getItem(key)](#getitemkey)  
  Gets a setting object by its key, which is case-sensitive. Throws an `ItemNotFound` error if the setting doesn't exist.
- [getItemOrNullObject(key)](#getitemornullobjectkey)  
  Gets a setting object by its key, which is case-sensitive. If the setting doesn't exist, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).
- [load(options)](#loadoptions)  
  Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- [load(propertyNames)](#loadpropertynames)  
  Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- [load(propertyNamesAndPaths)](#loadpropertynamesandpaths)  
  Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- [toJSON()](#tojson)  
  Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.SettingCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.SettingCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
- [track()](#track)  
  Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
- [untrack()](#untrack)  
  Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

---

## Property Details

### context
The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

- Property Value: [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### items
Gets the loaded child items in this collection.

```typescript
readonly items: Word.Setting[];
```

- Property Value: [Word.Setting](/en-us/javascript/api/word/word.setting)[]

#### Examples
```TypeScript
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

## Method Details

### add(key, value)
Creates a new setting or sets an existing setting.

```typescript
add(key: string, value: any): Word.Setting;
```

- Parameters:
  - key: string  
    Required. The setting's key, which is case-sensitive.
  - value: any  
    Required. The setting's value.
- Returns: [Word.Setting](/en-us/javascript/api/word/word.setting)

Remarks  
[ API set: WordApi 1.4 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples
```TypeScript
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

### deleteAll()
Deletes all settings in this add-in.

```typescript
deleteAll(): void;
```

- Returns: void

Remarks  
[ API set: WordApi 1.4 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples
```TypeScript
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

### getCount()
Gets the count of settings.

```typescript
getCount(): OfficeExtension.ClientResult<number>;
```

- Returns: [OfficeExtension.ClientResult](/en-us/javascript/api/office/officeextension.clientresult)<number>

Remarks  
[ API set: WordApi 1.4 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples
```TypeScript
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

### getItem(key)
Gets a setting object by its key, which is case-sensitive. Throws an `ItemNotFound` error if the setting doesn't exist.

```typescript
getItem(key: string): Word.Setting;
```

- Parameters:
  - key: string  
    The key that identifies the setting object.
- Returns: [Word.Setting](/en-us/javascript/api/word/word.setting)

Remarks  
[ API set: WordApi 1.4 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples
```TypeScript
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

### getItemOrNullObject(key)
Gets a setting object by its key, which is case-sensitive. If the setting doesn't exist, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
getItemOrNullObject(key: string): Word.Setting;
```

- Parameters:
  - key: string  
    Required. The key that identifies the setting object.
- Returns: [Word.Setting](/en-us/javascript/api/word/word.setting)

Remarks  
[ API set: WordApi 1.4 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples
```TypeScript
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

### load(options)
Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(options?: Word.Interfaces.SettingCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions): Word.SettingCollection;
```

- Parameters:
  - options: [Word.Interfaces.SettingCollectionLoadOptions](/en-us/javascript/api/word/word.interfaces.settingcollectionloadoptions) & [Word.Interfaces.CollectionLoadOptions](/en-us/javascript/api/word/word.interfaces.collectionloadoptions)  
    Provides options for which properties of the object to load.
- Returns: [Word.SettingCollection](/en-us/javascript/api/word/word.settingcollection)

### load(propertyNames)
Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.SettingCollection;
```

- Parameters:
  - propertyNames: string | string[]  
    A comma-delimited string or an array of strings that specify the properties to load.
- Returns: [Word.SettingCollection](/en-us/javascript/api/word/word.settingcollection)

### load(propertyNamesAndPaths)
Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Word.SettingCollection;
```

- Parameters:
  - propertyNamesAndPaths: [OfficeExtension.LoadOption](/en-us/javascript/api/office/officeextension.loadoption)  
    `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
- Returns: [Word.SettingCollection](/en-us/javascript/api/word/word.settingcollection)

### toJSON()
Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.SettingCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.SettingCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

```typescript
toJSON(): Word.Interfaces.SettingCollectionData;
```

- Returns: [Word.Interfaces.SettingCollectionData](/en-us/javascript/api/word/word.interfaces.settingcollectiondata)

### track()
Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.SettingCollection;
```

- Returns: [Word.SettingCollection](/en-us/javascript/api/word/word.settingcollection)

### untrack()
Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

```typescript
untrack(): Word.SettingCollection;
```

- Returns: [Word.SettingCollection](/en-us/javascript/api/word/word.settingcollection)