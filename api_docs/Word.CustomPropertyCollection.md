# Word.CustomPropertyCollection class

Package: [word](/en-us/javascript/api/word)

Contains the collection of [Word.CustomProperty](/en-us/javascript/api/word/word.customproperty) objects.

Extends: [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks
[API set: WordApi 1.3]

### Examples
```ts
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/30-properties/read-write-custom-document-properties.yaml

await Word.run(async (context) => {
    const properties: Word.CustomPropertyCollection = context.document.properties.customProperties;
    properties.load("key,type,value");

    await context.sync();
    for (let i = 0; i < properties.items.length; i++)
        console.log("Property Name:" + properties.items[i].key + "; Type=" + properties.items[i].type + "; Property Value=" + properties.items[i].value);
});
```

## Properties
- context — The request context associated with the object. This connects the add-in's process to the Office host application's process.
- items — Gets the loaded child items in this collection.

## Methods
- add(key, value) — Creates a new or sets an existing custom property.
- deleteAll() — Deletes all custom properties in this collection.
- getCount() — Gets the count of custom properties.
- getItem(key) — Gets a custom property object by its key, which is case-insensitive. Throws an ItemNotFound error if the custom property doesn't exist.
- getItemOrNullObject(key) — Gets a custom property object by its key, which is case-insensitive. If the custom property doesn't exist, returns an object with isNullObject set to true. See [OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).
- load(options) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNames) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNamesAndPaths) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- toJSON() — Overrides JavaScript toJSON() to provide useful output for JSON.stringify(). Returns a plain object (typed as Word.Interfaces.CustomPropertyCollectionData) with an "items" array of shallow copies of any loaded properties.
- track() — Track the object for automatic adjustment based on surrounding changes in the document. Shorthand for context.trackedObjects.add(thisObject).
- untrack() — Release the memory associated with this object, if it has previously been tracked. Shorthand for context.trackedObjects.remove(thisObject). Call context.sync() for the release to take effect.

## Property Details

### context
The request context associated with the object. This connects the add-in's process to the Office host application's process.

```ts
context: RequestContext;
```

Property Value: [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### items
Gets the loaded child items in this collection.

```ts
readonly items: Word.CustomProperty[];
```

Property Value: [Word.CustomProperty](/en-us/javascript/api/word/word.customproperty)[]

#### Examples
```ts
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/30-properties/read-write-custom-document-properties.yaml

await Word.run(async (context) => {
    const properties: Word.CustomPropertyCollection = context.document.properties.customProperties;
    properties.load("key,type,value");

    await context.sync();
    for (let i = 0; i < properties.items.length; i++)
        console.log("Property Name:" + properties.items[i].key + "; Type=" + properties.items[i].type + "; Property Value=" + properties.items[i].value);
});
```

## Method Details

### add(key, value)
Creates a new or sets an existing custom property.

```ts
add(key: string, value: any): Word.CustomProperty;
```

Parameters:
- key (string) — Required. The custom property's key, which is case-insensitive.
- value (any) — Required. The custom property's value.

Returns: [Word.CustomProperty](/en-us/javascript/api/word/word.customproperty)

Remarks:
[API set: WordApi 1.3]

#### Examples
```ts
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/30-properties/read-write-custom-document-properties.yaml

await Word.run(async (context) => {
    context.document.properties.customProperties.add("Numeric Property", 1234);

    await context.sync();
    console.log("Property added");
});

...

await Word.run(async (context) => {
    context.document.properties.customProperties.add("String Property", "Hello World!");

    await context.sync();
    console.log("Property added");
});
```

### deleteAll()
Deletes all custom properties in this collection.

```ts
deleteAll(): void;
```

Returns: void

Remarks:
[API set: WordApi 1.3]

### getCount()
Gets the count of custom properties.

```ts
getCount(): OfficeExtension.ClientResult<number>;
```

Returns: [OfficeExtension.ClientResult](/en-us/javascript/api/office/officeextension.clientresult)<number>

Remarks:
[API set: WordApi 1.3]

### getItem(key)
Gets a custom property object by its key, which is case-insensitive. Throws an ItemNotFound error if the custom property doesn't exist.

```ts
getItem(key: string): Word.CustomProperty;
```

Parameters:
- key (string) — The key that identifies the custom property object.

Returns: [Word.CustomProperty](/en-us/javascript/api/word/word.customproperty)

Remarks:
[API set: WordApi 1.3]

### getItemOrNullObject(key)
Gets a custom property object by its key, which is case-insensitive. If the custom property doesn't exist, then this method will return an object with its isNullObject property set to true. For further information, see [OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```ts
getItemOrNullObject(key: string): Word.CustomProperty;
```

Parameters:
- key (string) — Required. The key that identifies the custom property object.

Returns: [Word.CustomProperty](/en-us/javascript/api/word/word.customproperty)

Remarks:
[API set: WordApi 1.3]

### load(options)
Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```ts
load(options?: Word.Interfaces.CustomPropertyCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions): Word.CustomPropertyCollection;
```

Parameters:
- options ([Word.Interfaces.CustomPropertyCollectionLoadOptions](/en-us/javascript/api/word/word.interfaces.custompropertycollectionloadoptions) & [Word.Interfaces.CollectionLoadOptions](/en-us/javascript/api/word/word.interfaces.collectionloadoptions)) — Provides options for which properties of the object to load.

Returns: [Word.CustomPropertyCollection](/en-us/javascript/api/word/word.custompropertycollection)

### load(propertyNames)
Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```ts
load(propertyNames?: string | string[]): Word.CustomPropertyCollection;
```

Parameters:
- propertyNames (string | string[]) — A comma-delimited string or an array of strings that specify the properties to load.

Returns: [Word.CustomPropertyCollection](/en-us/javascript/api/word/word.custompropertycollection)

### load(propertyNamesAndPaths)
Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```ts
load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Word.CustomPropertyCollection;
```

Parameters:
- propertyNamesAndPaths ([OfficeExtension.LoadOption](/en-us/javascript/api/office/officeextension.loadoption)) — propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

Returns: [Word.CustomPropertyCollection](/en-us/javascript/api/word/word.custompropertycollection)

### toJSON()
Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.CustomPropertyCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.CustomPropertyCollectionData) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

```ts
toJSON(): Word.Interfaces.CustomPropertyCollectionData;
```

Returns: [Word.Interfaces.CustomPropertyCollectionData](/en-us/javascript/api/word/word.interfaces.custompropertycollectiondata)

### track()
Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```ts
track(): Word.CustomPropertyCollection;
```

Returns: [Word.CustomPropertyCollection](/en-us/javascript/api/word/word.custompropertycollection)

### untrack()
Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```ts
untrack(): Word.CustomPropertyCollection;
```

Returns: [Word.CustomPropertyCollection](/en-us/javascript/api/word/word.custompropertycollection)