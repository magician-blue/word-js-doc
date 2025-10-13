# Word.ListLevelCollection class

Package: [word](/en-us/javascript/api/word)

Contains a collection of [Word.ListLevel](/en-us/javascript/api/word/word.listlevel) objects.

Extends
- [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks
[ API set: WordApiDesktop 1.1 ]

#### Examples
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
- context  
  The request context associated with the object. This connects the add-in's process to the Office host application's process.
- items  
  Gets the loaded child items in this collection.

## Methods
- getFirst()  
  Gets the first list level in this collection. Throws an ItemNotFound error if this collection is empty.
- getFirstOrNullObject()  
  Gets the first list level in this collection. If this collection is empty, then this method will return an object with its isNullObject property set to true. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).
- load(options)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNames)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNamesAndPaths)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- toJSON()  
  Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.ListLevelCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.ListLevelCollectionData) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
- track()  
  Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
- untrack()  
  Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

## Property Details

### context
The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value  
[Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### items
Gets the loaded child items in this collection.

```typescript
readonly items: Word.ListLevel[];
```

Property Value  
[Word.ListLevel](/en-us/javascript/api/word/word.listlevel)[]

## Method Details

### getFirst()
Gets the first list level in this collection. Throws an ItemNotFound error if this collection is empty.

```typescript
getFirst(): Word.ListLevel;
```

Returns  
[Word.ListLevel](/en-us/javascript/api/word/word.listlevel)

Remarks  
[ API set: WordApiDesktop 1.1 ]

### getFirstOrNullObject()
Gets the first list level in this collection. If this collection is empty, then this method will return an object with its isNullObject property set to true. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
getFirstOrNullObject(): Word.ListLevel;
```

Returns  
[Word.ListLevel](/en-us/javascript/api/word/word.listlevel)

Remarks  
[ API set: WordApiDesktop 1.1 ]

### load(options)
Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(options?: Word.Interfaces.ListLevelCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions): Word.ListLevelCollection;
```

Parameters  
- options: [Word.Interfaces.ListLevelCollectionLoadOptions](/en-us/javascript/api/word/word.interfaces.listlevelcollectionloadoptions) & [Word.Interfaces.CollectionLoadOptions](/en-us/javascript/api/word/word.interfaces.collectionloadoptions)  
  Provides options for which properties of the object to load.

Returns  
[Word.ListLevelCollection](/en-us/javascript/api/word/word.listlevelcollection)

### load(propertyNames)
Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.ListLevelCollection;
```

Parameters  
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns  
[Word.ListLevelCollection](/en-us/javascript/api/word/word.listlevelcollection)

### load(propertyNamesAndPaths)
Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Word.ListLevelCollection;
```

Parameters  
- propertyNamesAndPaths: [OfficeExtension.LoadOption](/en-us/javascript/api/office/officeextension.loadoption)  
  propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

Returns  
[Word.ListLevelCollection](/en-us/javascript/api/word/word.listlevelcollection)

### toJSON()
Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.ListLevelCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.ListLevelCollectionData) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

```typescript
toJSON(): Word.Interfaces.ListLevelCollectionData;
```

Returns  
[Word.Interfaces.ListLevelCollectionData](/en-us/javascript/api/word/word.interfaces.listlevelcollectiondata)

### track()
Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.ListLevelCollection;
```

Returns  
[Word.ListLevelCollection](/en-us/javascript/api/word/word.listlevelcollection)

### untrack()
Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.ListLevelCollection;
```

Returns  
[Word.ListLevelCollection](/en-us/javascript/api/word/word.listlevelcollection)