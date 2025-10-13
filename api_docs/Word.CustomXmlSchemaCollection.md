# Word.CustomXmlSchemaCollection class

Package: [word](https://learn.microsoft.com/en-us/javascript/api/word)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents a collection of [Word.CustomXmlSchema](https://learn.microsoft.com/en-us/javascript/api/word/word.customxmlschema) objects attached to a data stream.

Extends: [OfficeExtension.ClientObject](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientobject)

## Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties
- context: The request context associated with the object. This connects the add-in's process to the Office host application's process.
- items: Gets the loaded child items in this collection.

## Methods
- add(options): Adds one or more schemas to the schema collection that can then be added to a stream in the data store and to the schema library.
- addCollection(schemaCollection): Adds an existing schema collection to the current schema collection.
- getCount(): Returns the number of items in the collection.
- getItem(index): Returns a CustomXmlSchema object that represents the specified item in the collection.
- getNamespaceUri(): Returns the number of items in the collection.
- load(options): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNames): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNamesAndPaths): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- toJSON(): Overrides the JavaScript toJSON() method to provide more useful output when an API object is passed to JSON.stringify().
- track(): Track the object for automatic adjustment based on surrounding changes in the document.
- untrack(): Release the memory associated with this object, if it has previously been tracked.
- validate(): Specifies whether the schemas in the schema collection are valid (conforms to the syntactic rules of XML and the rules for a specified vocabulary).

## Property Details

### context
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value: [Word.RequestContext](https://learn.microsoft.com/en-us/javascript/api/word/word.requestcontext)

### items
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the loaded child items in this collection.

```typescript
readonly items: Word.CustomXmlSchema[];
```

Property Value: [Word.CustomXmlSchema](https://learn.microsoft.com/en-us/javascript/api/word/word.customxmlschema)[]

## Method Details

### add(options)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Adds one or more schemas to the schema collection that can then be added to a stream in the data store and to the schema library.

```typescript
add(options?: Word.CustomXmlAddSchemaOptions): Word.CustomXmlSchema;
```

Parameters:
- options: [Word.CustomXmlAddSchemaOptions](https://learn.microsoft.com/en-us/javascript/api/word/word.customxmladdschemaoptions)  
  Optional. The options that define the schema to be added.

Returns: [Word.CustomXmlSchema](https://learn.microsoft.com/en-us/javascript/api/word/word.customxmlschema)

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### addCollection(schemaCollection)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Adds an existing schema collection to the current schema collection.

```typescript
addCollection(schemaCollection: Word.CustomXmlSchemaCollection): Word.CustomXmlSchemaCollection;
```

Parameters:
- schemaCollection: [Word.CustomXmlSchemaCollection](https://learn.microsoft.com/en-us/javascript/api/word/word.customxmlschemacollection)  
  The schema collection to add.

Returns: [Word.CustomXmlSchemaCollection](https://learn.microsoft.com/en-us/javascript/api/word/word.customxmlschemacollection)

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### getCount()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the number of items in the collection.

```typescript
getCount(): OfficeExtension.ClientResult<number>;
```

Returns: [OfficeExtension.ClientResult](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientresult)<number>

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### getItem(index)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `CustomXmlSchema` object that represents the specified item in the collection.

```typescript
getItem(index: number): Word.CustomXmlSchema;
```

Parameters:
- index: number  
  A number that identifies the index location of a paragraph object.

Returns: [Word.CustomXmlSchema](https://learn.microsoft.com/en-us/javascript/api/word/word.customxmlschema)

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### getNamespaceUri()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the number of items in the collection.

```typescript
getNamespaceUri(): OfficeExtension.ClientResult<number>;
```

Returns: [OfficeExtension.ClientResult](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientresult)<number>

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### load(options)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(options?: Word.Interfaces.CustomXmlSchemaCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions): Word.CustomXmlSchemaCollection;
```

Parameters:
- options: [Word.Interfaces.CustomXmlSchemaCollectionLoadOptions](https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.customxmlschemacollectionloadoptions) & [Word.Interfaces.CollectionLoadOptions](https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.collectionloadoptions)  
  Provides options for which properties of the object to load.

Returns: [Word.CustomXmlSchemaCollection](https://learn.microsoft.com/en-us/javascript/api/word/word.customxmlschemacollection)

### load(propertyNames)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.CustomXmlSchemaCollection;
```

Parameters:
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns: [Word.CustomXmlSchemaCollection](https://learn.microsoft.com/en-us/javascript/api/word/word.customxmlschemacollection)

### load(propertyNamesAndPaths)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Word.CustomXmlSchemaCollection;
```

Parameters:
- propertyNamesAndPaths: [OfficeExtension.LoadOption](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.loadoption)  
  `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

Returns: [Word.CustomXmlSchemaCollection](https://learn.microsoft.com/en-us/javascript/api/word/word.customxmlschemacollection)

### toJSON()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.CustomXmlSchemaCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.CustomXmlSchemaCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

```typescript
toJSON(): Word.Interfaces.CustomXmlSchemaCollectionData;
```

Returns: [Word.Interfaces.CustomXmlSchemaCollectionData](https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.customxmlschemacollectiondata)

### track()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.CustomXmlSchemaCollection;
```

Returns: [Word.CustomXmlSchemaCollection](https://learn.microsoft.com/en-us/javascript/api/word/word.customxmlschemacollection)

### untrack()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

```typescript
untrack(): Word.CustomXmlSchemaCollection;
```

Returns: [Word.CustomXmlSchemaCollection](https://learn.microsoft.com/en-us/javascript/api/word/word.customxmlschemacollection)

### validate()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the schemas in the schema collection are valid (conforms to the syntactic rules of XML and the rules for a specified vocabulary).

```typescript
validate(): OfficeExtension.ClientResult<boolean>;
```

Returns: [OfficeExtension.ClientResult](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientresult)<boolean>

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)