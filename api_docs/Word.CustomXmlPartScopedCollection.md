# Word.CustomXmlPartScopedCollection class

- Package: [word](/en-us/javascript/api/word)

Contains the collection of [Word.CustomXmlPart](/en-us/javascript/api/word/word.customxmlpart) objects with a specific namespace.

- Extends: [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks
[ API set: WordApi 1.4 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples
```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-custom-xml-part-ns.yaml

// Original XML: <Reviewers xmlns='http://schemas.contoso.com/review/1.0'><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>

// Gets the custom XML parts with the specified namespace URI.
await Word.run(async (context) => {
  const namespaceUri = "http://schemas.contoso.com/review/1.0";
  console.log(`Specified namespace URI: ${namespaceUri}`);
  const scopedCustomXmlParts: Word.CustomXmlPartScopedCollection =
    context.document.customXmlParts.getByNamespace(namespaceUri);
  scopedCustomXmlParts.load("items");
  await context.sync();

  console.log(`Number of custom XML parts found with this namespace: ${!scopedCustomXmlParts.items ? 0 : scopedCustomXmlParts.items.length}`);
});
```

## Properties
- context — The request context associated with the object. This connects the add-in's process to the Office host application's process.
- items — Gets the loaded child items in this collection.

## Methods
- getCount() — Gets the number of items in the collection.
- getItem(id) — Gets a custom XML part based on its ID.
- getItemOrNullObject(id) — Gets a custom XML part based on its ID. If the CustomXmlPart doesn't exist in the collection, then this method will return an object with its isNullObject property set to true. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).
- getOnlyItem() — If the collection contains exactly one item, this method returns it. Otherwise, this method produces an error.
- getOnlyItemOrNullObject() — If the collection contains exactly one item, this method returns it. Otherwise, this method will return an object with its isNullObject property set to true. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).
- load(options) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNames) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNamesAndPaths) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- toJSON() — Overrides the JavaScript toJSON() method to provide more useful output when an API object is passed to JSON.stringify(). Returns a plain JavaScript object (typed as Word.Interfaces.CustomXmlPartScopedCollectionData) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
- track() — Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
- untrack() — Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

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
readonly items: Word.CustomXmlPart[];
```

Property value: [Word.CustomXmlPart](/en-us/javascript/api/word/word.customxmlpart)[]

## Method Details

### getCount()
Gets the number of items in the collection.

```typescript
getCount(): OfficeExtension.ClientResult<number>;
```

Returns: [OfficeExtension.ClientResult](/en-us/javascript/api/office/officeextension.clientresult)<number>

Remarks: [ API set: WordApi 1.4 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### getItem(id)
Gets a custom XML part based on its ID.

```typescript
getItem(id: string): Word.CustomXmlPart;
```

Parameters:
- id: string — ID of the custom XML part to be retrieved.

Returns: [Word.CustomXmlPart](/en-us/javascript/api/word/word.customxmlpart)

Remarks: [ API set: WordApi 1.4 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### getItemOrNullObject(id)
Gets a custom XML part based on its ID. If the CustomXmlPart doesn't exist in the collection, then this method will return an object with its isNullObject property set to true. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
getItemOrNullObject(id: string): Word.CustomXmlPart;
```

Parameters:
- id: string — Required. ID of the object to be retrieved.

Returns: [Word.CustomXmlPart](/en-us/javascript/api/word/word.customxmlpart)

Remarks: [ API set: WordApi 1.4 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### getOnlyItem()
If the collection contains exactly one item, this method returns it. Otherwise, this method produces an error.

```typescript
getOnlyItem(): Word.CustomXmlPart;
```

Returns: [Word.CustomXmlPart](/en-us/javascript/api/word/word.customxmlpart)

Remarks: [ API set: WordApi 1.4 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### getOnlyItemOrNullObject()
If the collection contains exactly one item, this method returns it. Otherwise, this method will return an object with its isNullObject property set to true. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
getOnlyItemOrNullObject(): Word.CustomXmlPart;
```

Returns: [Word.CustomXmlPart](/en-us/javascript/api/word/word.customxmlpart)

Remarks: [ API set: WordApi 1.4 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### load(options)
Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(options?: Word.Interfaces.CustomXmlPartScopedCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions): Word.CustomXmlPartScopedCollection;
```

Parameters:
- options: [Word.Interfaces.CustomXmlPartScopedCollectionLoadOptions](/en-us/javascript/api/word/word.interfaces.customxmlpartscopedcollectionloadoptions) & [Word.Interfaces.CollectionLoadOptions](/en-us/javascript/api/word/word.interfaces.collectionloadoptions) — Provides options for which properties of the object to load.

Returns: [Word.CustomXmlPartScopedCollection](/en-us/javascript/api/word/word.customxmlpartscopedcollection)

### load(propertyNames)
Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.CustomXmlPartScopedCollection;
```

Parameters:
- propertyNames: string | string[] — A comma-delimited string or an array of strings that specify the properties to load.

Returns: [Word.CustomXmlPartScopedCollection](/en-us/javascript/api/word/word.customxmlpartscopedcollection)

### load(propertyNamesAndPaths)
Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Word.CustomXmlPartScopedCollection;
```

Parameters:
- propertyNamesAndPaths: [OfficeExtension.LoadOption](/en-us/javascript/api/office/officeextension.loadoption) — propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

Returns: [Word.CustomXmlPartScopedCollection](/en-us/javascript/api/word/word.customxmlpartscopedcollection)

### toJSON()
Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.CustomXmlPartScopedCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.CustomXmlPartScopedCollectionData) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

```typescript
toJSON(): Word.Interfaces.CustomXmlPartScopedCollectionData;
```

Returns: [Word.Interfaces.CustomXmlPartScopedCollectionData](/en-us/javascript/api/word/word.interfaces.customxmlpartscopedcollectiondata)

### track()
Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.CustomXmlPartScopedCollection;
```

Returns: [Word.CustomXmlPartScopedCollection](/en-us/javascript/api/word/word.customxmlpartscopedcollection)

### untrack()
Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.CustomXmlPartScopedCollection;
```

Returns: [Word.CustomXmlPartScopedCollection](/en-us/javascript/api/word/word.customxmlpartscopedcollection)