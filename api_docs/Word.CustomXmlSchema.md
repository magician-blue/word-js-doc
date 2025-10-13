# Word.CustomXmlSchema class

Package: [word](/en-us/javascript/api/word)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents a schema in a [Word.CustomXmlSchemaCollection](/en-us/javascript/api/word/word.customxmlschemacollection) object.

- Extends: [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

## Properties
- context  
  The request context associated with the object. This connects the add-in's process to the Office host application's process.

- location  
  Gets the location of the schema on a computer.

- namespaceUri  
  Gets the unique address identifier for the namespace of the CustomXmlSchema object.

## Methods
- delete()  
  Deletes this schema from the [Word.CustomXmlSchemaCollection](/en-us/javascript/api/word/word.customxmlschemacollection) object.

- load(options)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

- load(propertyNames)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

- load(propertyNamesAndPaths)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

- reload()  
  Reloads the schema from a file.

- toJSON()  
  Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.CustomXmlSchema object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.CustomXmlSchemaData) that contains shallow copies of any loaded child properties from the original object.

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

- Type: [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### location
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the location of the schema on a computer.

```typescript
readonly location: string;
```

- Type: string

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

### namespaceUri
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the unique address identifier for the namespace of the CustomXmlSchema object.

```typescript
readonly namespaceUri: string;
```

- Type: string

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

## Method Details

### delete()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Deletes this schema from the [Word.CustomXmlSchemaCollection](/en-us/javascript/api/word/word.customxmlschemacollection) object.

```typescript
delete(): void;
```

- Returns: void

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

### load(options)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(options?: Word.Interfaces.CustomXmlSchemaLoadOptions): Word.CustomXmlSchema;
```

- Parameters:
  - options: [Word.Interfaces.CustomXmlSchemaLoadOptions](/en-us/javascript/api/word/word.interfaces.customxmlschemaloadoptions)  
    Provides options for which properties of the object to load.
- Returns: [Word.CustomXmlSchema](/en-us/javascript/api/word/word.customxmlschema)

### load(propertyNames)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.CustomXmlSchema;
```

- Parameters:
  - propertyNames: string | string[]  
    A comma-delimited string or an array of strings that specify the properties to load.
- Returns: [Word.CustomXmlSchema](/en-us/javascript/api/word/word.customxmlschema)

### load(propertyNamesAndPaths)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
    select?: string;
    expand?: string;
}): Word.CustomXmlSchema;
```

- Parameters:
  - propertyNamesAndPaths:  
    - select?: string  
    - expand?: string  
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.
- Returns: [Word.CustomXmlSchema](/en-us/javascript/api/word/word.customxmlschema)

### reload()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Reloads the schema from a file.

```typescript
reload(): void;
```

- Returns: void

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

### toJSON()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.CustomXmlSchema object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.CustomXmlSchemaData) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.CustomXmlSchemaData;
```

- Returns: [Word.Interfaces.CustomXmlSchemaData](/en-us/javascript/api/word/word.interfaces.customxmlschemadata)

### track()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.CustomXmlSchema;
```

- Returns: [Word.CustomXmlSchema](/en-us/javascript/api/word/word.customxmlschema)

### untrack()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.CustomXmlSchema;
```

- Returns: [Word.CustomXmlSchema](/en-us/javascript/api/word/word.customxmlschema)