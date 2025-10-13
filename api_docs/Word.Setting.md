# Word.Setting class

Package: [word](/en-us/javascript/api/word)

Represents a setting of the add-in.

Extends: [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks

[API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples

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

- context  
  The request context associated with the object. This connects the add-in's process to the Office host application's process.

- key  
  Gets the key of the setting.

- value  
  Specifies the value of the setting.

## Methods

- delete()  
  Deletes the setting.

- load(options)  
  Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

- load(propertyNames)  
  Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

- load(propertyNamesAndPaths)  
  Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

- set(properties, options)  
  Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

- set(properties)  
  Sets multiple properties on the object at the same time, based on an existing loaded object.

- toJSON()  
  Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.Setting` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.SettingData`) that contains shallow copies of any loaded child properties from the original object.

- track()  
  Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

- untrack()  
  Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

## Property Details

### context

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value: [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### key

Gets the key of the setting.

```typescript
readonly key: string;
```

Property Value: string

#### Remarks

[API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples

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

### value

Specifies the value of the setting.

```typescript
value: any;
```

Property Value: any

#### Remarks

[API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples

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

## Method Details

### delete()

Deletes the setting.

```typescript
delete(): void;
```

Returns: void

#### Remarks

[API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples

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

### load(options)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(options?: Word.Interfaces.SettingLoadOptions): Word.Setting;
```

Parameters:

- options: [Word.Interfaces.SettingLoadOptions](/en-us/javascript/api/word/word.interfaces.settingloadoptions)  
  Provides options for which properties of the object to load.

Returns: [Word.Setting](/en-us/javascript/api/word/word.setting)

### load(propertyNames)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.Setting;
```

Parameters:

- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns: [Word.Setting](/en-us/javascript/api/word/word.setting)

### load(propertyNamesAndPaths)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.Setting;
```

Parameters:

- propertyNamesAndPaths:  
  {
  select?: string;
  expand?: string;
  }

  `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

Returns: [Word.Setting](/en-us/javascript/api/word/word.setting)

### set(properties, options)

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.SettingUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

Parameters:

- properties: [Word.Interfaces.SettingUpdateData](/en-us/javascript/api/word/word.interfaces.settingupdatedata)  
  A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.

- options: [OfficeExtension.UpdateOptions](/en-us/javascript/api/office/officeextension.updateoptions)  
  Provides an option to suppress errors if the properties object tries to set any read-only properties.

Returns: void

### set(properties)

Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.Setting): void;
```

Parameters:

- properties: [Word.Setting](/en-us/javascript/api/word/word.setting)

Returns: void

### toJSON()

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.Setting` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.SettingData`) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.SettingData;
```

Returns: [Word.Interfaces.SettingData](/en-us/javascript/api/word/word.interfaces.settingdata)

### track()

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.Setting;
```

Returns: [Word.Setting](/en-us/javascript/api/word/word.setting)

### untrack()

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

```typescript
untrack(): Word.Setting;
```

Returns: [Word.Setting](/en-us/javascript/api/word/word.setting)