# Word.CustomXmlPrefixMappingCollection class

Package: [word](/en-us/javascript/api/word)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents a collection of [Word.CustomXmlPrefixMapping](/en-us/javascript/api/word/word.customxmlprefixmapping) objects.

Extends: [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- context — The request context associated with the object. This connects the add-in's process to the Office host application's process.
- items — Gets the loaded child items in this collection.

## Methods

- addNamespace(prefix, namespaceUri) — Adds a custom namespace/prefix mapping to use when querying an item.
- getCount() — Returns the number of items in the collection.
- getItem(index) — Returns a CustomXmlPrefixMapping object that represents the specified item in the collection.
- load(options) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNames) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNamesAndPaths) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- lookupNamespace(prefix) — Gets the namespace corresponding to the specified prefix.
- lookupPrefix(namespaceUri) — Gets the prefix corresponding to the specified namespace.
- toJSON() — Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify().
- track() — Track the object for automatic adjustment based on surrounding changes in the document.
- untrack() — Release the memory associated with this object, if it has previously been tracked.

## Property Details

### context

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value: [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### items

Gets the loaded child items in this collection.

```typescript
readonly items: Word.CustomXmlPrefixMapping[];
```

Property Value: [Word.CustomXmlPrefixMapping](/en-us/javascript/api/word/word.customxmlprefixmapping)[]

## Method Details

### addNamespace(prefix, namespaceUri)

Adds a custom namespace/prefix mapping to use when querying an item.

```typescript
addNamespace(prefix: string, namespaceUri: string): OfficeExtension.ClientResult<number>;
```

Parameters:
- prefix: string  
  The prefix to associate with the namespace.
- namespaceUri: string  
  The namespace URI to map.

Returns: [OfficeExtension.ClientResult](/en-us/javascript/api/office/officeextension.clientresult)<number>

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### getCount()

Returns the number of items in the collection.

```typescript
getCount(): OfficeExtension.ClientResult<number>;
```

Returns: [OfficeExtension.ClientResult](/en-us/javascript/api/office/officeextension.clientresult)<number>

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### getItem(index)

Returns a `CustomXmlPrefixMapping` object that represents the specified item in the collection.

```typescript
getItem(index: number): Word.CustomXmlPrefixMapping;
```

Parameters:
- index: number  
  A number that identifies the index location of a paragraph object.

Returns: [Word.CustomXmlPrefixMapping](/en-us/javascript/api/word/word.customxmlprefixmapping)

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### load(options)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(options?: Word.Interfaces.CustomXmlPrefixMappingCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions): Word.CustomXmlPrefixMappingCollection;
```

Parameters:
- options: [Word.Interfaces.CustomXmlPrefixMappingCollectionLoadOptions](/en-us/javascript/api/word/word.interfaces.customxmlprefixmappingcollectionloadoptions) & [Word.Interfaces.CollectionLoadOptions](/en-us/javascript/api/word/word.interfaces.collectionloadoptions)  
  Provides options for which properties of the object to load.

Returns: [Word.CustomXmlPrefixMappingCollection](/en-us/javascript/api/word/word.customxmlprefixmappingcollection)

### load(propertyNames)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.CustomXmlPrefixMappingCollection;
```

Parameters:
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns: [Word.CustomXmlPrefixMappingCollection](/en-us/javascript/api/word/word.customxmlprefixmappingcollection)

### load(propertyNamesAndPaths)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Word.CustomXmlPrefixMappingCollection;
```

Parameters:
- propertyNamesAndPaths: [OfficeExtension.LoadOption](/en-us/javascript/api/office/officeextension.loadoption)  
  `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

Returns: [Word.CustomXmlPrefixMappingCollection](/en-us/javascript/api/word/word.customxmlprefixmappingcollection)

### lookupNamespace(prefix)

Gets the namespace corresponding to the specified prefix.

```typescript
lookupNamespace(prefix: string): OfficeExtension.ClientResult<string>;
```

Parameters:
- prefix: string  
  The prefix to look up.

Returns: [OfficeExtension.ClientResult](/en-us/javascript/api/office/officeextension.clientresult)<string>

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### lookupPrefix(namespaceUri)

Gets the prefix corresponding to the specified namespace.

```typescript
lookupPrefix(namespaceUri: string): OfficeExtension.ClientResult<string>;
```

Parameters:
- namespaceUri: string  
  The namespace URI to look up.

Returns: [OfficeExtension.ClientResult](/en-us/javascript/api/office/officeextension.clientresult)<string>

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### toJSON()

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.CustomXmlPrefixMappingCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.CustomXmlPrefixMappingCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

```typescript
toJSON(): Word.Interfaces.CustomXmlPrefixMappingCollectionData;
```

Returns: [Word.Interfaces.CustomXmlPrefixMappingCollectionData](/en-us/javascript/api/word/word.interfaces.customxmlprefixmappingcollectiondata)

### track()

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.CustomXmlPrefixMappingCollection;
```

Returns: [Word.CustomXmlPrefixMappingCollection](/en-us/javascript/api/word/word.customxmlprefixmappingcollection)

### untrack()

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

```typescript
untrack(): Word.CustomXmlPrefixMappingCollection;
```

Returns: [Word.CustomXmlPrefixMappingCollection](/en-us/javascript/api/word/word.customxmlprefixmappingcollection)