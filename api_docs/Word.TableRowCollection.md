# Word.TableRowCollection class

Package: [word](/en-us/javascript/api/word)

Contains the collection of the document's TableRow objects.

Extends: [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks

[ API set: WordApi 1.3 ]

#### Examples

```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/40-tables/manage-formatting.yaml

// Gets content alignment details about the first row of the first table in the document.
await Word.run(async (context) => {
  const firstTable: Word.Table = context.document.body.tables.getFirst();
  const firstTableRow: Word.TableRow = firstTable.rows.getFirst();
  firstTableRow.load(["horizontalAlignment", "verticalAlignment"]);
  await context.sync();

  console.log(
    `Details about the alignment of the first table's first row:`,
    `- Horizontal alignment of every cell in the row: ${firstTableRow.horizontalAlignment}`,
    `- Vertical alignment of every cell in the row: ${firstTableRow.verticalAlignment}`
  );
});
```

## Properties

- [context](#word-word-tablerowcollection-context-member)  
  The request context associated with the object. This connects the add-in's process to the Office host application's process.
- [items](#word-word-tablerowcollection-items-member)  
  Gets the loaded child items in this collection.

## Methods

- [getFirst()](#word-word-tablerowcollection-getfirst-member1)  
  Gets the first row in this collection. Throws an `ItemNotFound` error if this collection is empty.
- [getFirstOrNullObject()](#word-word-tablerowcollection-getfirstornullobject-member1)  
  Gets the first row in this collection. If this collection is empty, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).
- [load(options)](#word-word-tablerowcollection-load-member1)  
  Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- [load(propertyNames)](#word-word-tablerowcollection-load-member2)  
  Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- [load(propertyNamesAndPaths)](#word-word-tablerowcollection-load-member3)  
  Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- [toJSON()](#word-word-tablerowcollection-tojson-member1)  
  Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.TableRowCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.TableRowCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
- [track()](#word-word-tablerowcollection-track-member1)  
  Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
- [untrack()](#word-word-tablerowcollection-untrack-member1)  
  Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

## Property Details

### context

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property value: [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### items

Gets the loaded child items in this collection.

```typescript
readonly items: Word.TableRow[];
```

Property value: [Word.TableRow](/en-us/javascript/api/word/word.tablerow)[]

## Method Details

### getFirst()

Gets the first row in this collection. Throws an `ItemNotFound` error if this collection is empty.

```typescript
getFirst(): Word.TableRow;
```

Returns: [Word.TableRow](/en-us/javascript/api/word/word.tablerow)

Remarks  
[ API set: WordApi 1.3 ]

#### Examples

```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/40-tables/manage-formatting.yaml

// Gets border details about the first row of the first table in the document.
await Word.run(async (context) => {
  const firstTable: Word.Table = context.document.body.tables.getFirst();
  const firstTableRow: Word.TableRow = firstTable.rows.getFirst();
  const borderLocation = Word.BorderLocation.bottom;
  const border: Word.TableBorder = firstTableRow.getBorder(borderLocation);
  border.load(["type", "color", "width"]);
  await context.sync();

  console.log(
    `Details about the ${borderLocation} border of the first table's first row:`,
    `- Color: ${border.color}`,
    `- Type: ${border.type}`,
    `- Width: ${border.width} points`
  );
});
```

### getFirstOrNullObject()

Gets the first row in this collection. If this collection is empty, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
getFirstOrNullObject(): Word.TableRow;
```

Returns: [Word.TableRow](/en-us/javascript/api/word/word.tablerow)

Remarks  
[ API set: WordApi 1.3 ]

### load(options)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(options?: Word.Interfaces.TableRowCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions): Word.TableRowCollection;
```

Parameters

- options: [Word.Interfaces.TableRowCollectionLoadOptions](/en-us/javascript/api/word/word.interfaces.tablerowcollectionloadoptions) & [Word.Interfaces.CollectionLoadOptions](/en-us/javascript/api/word/word.interfaces.collectionloadoptions)  
  Provides options for which properties of the object to load.

Returns: [Word.TableRowCollection](/en-us/javascript/api/word/word.tablerowcollection)

### load(propertyNames)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.TableRowCollection;
```

Parameters

- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns: [Word.TableRowCollection](/en-us/javascript/api/word/word.tablerowcollection)

### load(propertyNamesAndPaths)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Word.TableRowCollection;
```

Parameters

- propertyNamesAndPaths: [OfficeExtension.LoadOption](/en-us/javascript/api/office/officeextension.loadoption)  
  `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

Returns: [Word.TableRowCollection](/en-us/javascript/api/word/word.tablerowcollection)

### toJSON()

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.TableRowCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.TableRowCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

```typescript
toJSON(): Word.Interfaces.TableRowCollectionData;
```

Returns: [Word.Interfaces.TableRowCollectionData](/en-us/javascript/api/word/word.interfaces.tablerowcollectiondata)

### track()

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.TableRowCollection;
```

Returns: [Word.TableRowCollection](/en-us/javascript/api/word/word.tablerowcollection)

### untrack()

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

```typescript
untrack(): Word.TableRowCollection;
```

Returns: [Word.TableRowCollection](/en-us/javascript/api/word/word.tablerowcollection)