# Word.FieldCollection class

Package: [word](/en-us/javascript/api/word)

Contains a collection of [Word.Field](/en-us/javascript/api/word/word.field) objects.

Extends
- [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks

[ API set: WordApi 1.4 ]

Important: To learn more about which fields can be inserted, see the `Word.Range.insertField` API introduced in requirement set 1.5. Support for managing fields is similar to what's available in the Word UI. However, the Word UI on the web primarily only supports fields as read-only (see [Field codes in Word for the web](https://support.microsoft.com/office/d8f46094-13c3-4966-98c3-259748f3caf1)). To learn more about Word UI clients that more fully support fields, see the product list at the beginning of [Insert, edit, and view fields in Word](https://support.microsoft.com/office/c429bbb0-8669-48a7-bd24-bab6ba6b06bb).

#### Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-fields.yaml

// Gets all fields in the document body.
await Word.run(async (context) => {
  const fields: Word.FieldCollection = context.document.body.fields.load("items");

  await context.sync();

  if (fields.items.length === 0) {
    console.log("No fields in this document.");
  } else {
    fields.load(["code", "result"]);
    await context.sync();

    for (let i = 0; i < fields.items.length; i++) {
      console.log(`Field ${i + 1}'s code: ${fields.items[i].code}`, `Field ${i + 1}'s result: ${JSON.stringify(fields.items[i].result)}`);
    }
  }
});
```

## Properties

- [context](#word-word-fieldcollection-context-member)
  - The request context associated with the object. This connects the add-in's process to the Office host application's process.
- [items](#word-word-fieldcollection-items-member)
  - Gets the loaded child items in this collection.

## Methods

- [getByTypes(types)](#word-word-fieldcollection-getbytypes-member1)
  - Gets the Field object collection including the specified types of fields.
- [getFirst()](#word-word-fieldcollection-getfirst-member1)
  - Gets the first field in this collection. Throws an `ItemNotFound` error if this collection is empty.
- [getFirstOrNullObject()](#word-word-fieldcollection-getfirstornullobject-member1)
  - Gets the first field in this collection. If this collection is empty, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).
- [load(options)](#word-word-fieldcollection-load-member1)
  - Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- [load(propertyNames)](#word-word-fieldcollection-load-member2)
  - Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- [load(propertyNamesAndPaths)](#word-word-fieldcollection-load-member3)
  - Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- [toJSON()](#word-word-fieldcollection-tojson-member1)
  - Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.FieldCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.FieldCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
- [track()](#word-word-fieldcollection-track-member1)
  - Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
- [untrack()](#word-word-fieldcollection-untrack-member1)
  - Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

## Property Details

### context

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value
- [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### items

Gets the loaded child items in this collection.

```typescript
readonly items: Word.Field[];
```

Property Value
- [Word.Field](/en-us/javascript/api/word/word.field)[]

#### Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-fields.yaml

// Gets all fields in the document body.
await Word.run(async (context) => {
  const fields: Word.FieldCollection = context.document.body.fields.load("items");

  await context.sync();

  if (fields.items.length === 0) {
    console.log("No fields in this document.");
  } else {
    fields.load(["code", "result"]);
    await context.sync();

    for (let i = 0; i < fields.items.length; i++) {
      console.log(`Field ${i + 1}'s code: ${fields.items[i].code}`, `Field ${i + 1}'s result: ${JSON.stringify(fields.items[i].result)}`);
    }
  }
});
```

## Method Details

### getByTypes(types)

Gets the Field object collection including the specified types of fields.

```typescript
getByTypes(types: Word.FieldType[]): Word.FieldCollection;
```

Parameters
- types: [Word.FieldType](/en-us/javascript/api/word/word.fieldtype)[]
  - Required. An array of field types.

Returns
- [Word.FieldCollection](/en-us/javascript/api/word/word.fieldcollection)

Remarks
- [ API set: WordApi 1.5 ]

### getFirst()

Gets the first field in this collection. Throws an `ItemNotFound` error if this collection is empty.

```typescript
getFirst(): Word.Field;
```

Returns
- [Word.Field](/en-us/javascript/api/word/word.field)

Remarks
- [ API set: WordApi 1.4 ]

### getFirstOrNullObject()

Gets the first field in this collection. If this collection is empty, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
getFirstOrNullObject(): Word.Field;
```

Returns
- [Word.Field](/en-us/javascript/api/word/word.field)

Remarks
- [ API set: WordApi 1.4 ]

#### Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-fields.yaml

// Gets the first field in the document.
await Word.run(async (context) => {
  const field: Word.Field = context.document.body.fields.getFirstOrNullObject();
  field.load(["code", "result", "locked", "type", "data", "kind"]);

  await context.sync();

  if (field.isNullObject) {
    console.log("This document has no fields.");
  } else {
    console.log(
      "Code of first field: " + field.code,
      "Result of first field: " + JSON.stringify(field.result),
      "Type of first field: " + field.type,
      "Is the first field locked? " + field.locked,
      "Kind of the first field: " + field.kind
    );
  }
});
```

### load(options)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(options?: Word.Interfaces.FieldCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions): Word.FieldCollection;
```

Parameters
- options: [Word.Interfaces.FieldCollectionLoadOptions](/en-us/javascript/api/word/word.interfaces.fieldcollectionloadoptions) & [Word.Interfaces.CollectionLoadOptions](/en-us/javascript/api/word/word.interfaces.collectionloadoptions)  
  Provides options for which properties of the object to load.

Returns
- [Word.FieldCollection](/en-us/javascript/api/word/word.fieldcollection)

### load(propertyNames)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.FieldCollection;
```

Parameters
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns
- [Word.FieldCollection](/en-us/javascript/api/word/word.fieldcollection)

### load(propertyNamesAndPaths)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Word.FieldCollection;
```

Parameters
- propertyNamesAndPaths: [OfficeExtension.LoadOption](/en-us/javascript/api/office/officeextension.loadoption)  
  `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

Returns
- [Word.FieldCollection](/en-us/javascript/api/word/word.fieldcollection)

### toJSON()

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.FieldCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.FieldCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

```typescript
toJSON(): Word.Interfaces.FieldCollectionData;
```

Returns
- [Word.Interfaces.FieldCollectionData](/en-us/javascript/api/word/word.interfaces.fieldcollectiondata)

### track()

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.FieldCollection;
```

Returns
- [Word.FieldCollection](/en-us/javascript/api/word/word.fieldcollection)

### untrack()

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

```typescript
untrack(): Word.FieldCollection;
```

Returns
- [Word.FieldCollection](/en-us/javascript/api/word/word.fieldcollection)