# Word.TemplateCollection class

Package: [word](/en-us/javascript/api/word)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Contains a collection of [Word.Template](/en-us/javascript/api/word/word.template) objects that represent all the templates that are currently available. This collection includes open templates, templates attached to open documents, and global templates loaded in the Templates and Add-ins dialog box. To learn how to access this dialog in the Word UI, see Load or unload a template or add-in program: https://support.microsoft.com/office/2479fe53-f849-4394-88bb-2a6e2a39479d.

Extends
- [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- context  
  The request context associated with the object. This connects the add-in's process to the Office host application's process.

- items  
  Gets the loaded child items in this collection.

## Methods

- getCount()  
  Returns the number of items in the collection.

- getItemAt(index)  
  Gets a Template object by its index in the collection.

- importBuildingBlocks()  
  Imports the building blocks for all templates into Microsoft Word.

- load(options)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

- load(propertyNames)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

- load(propertyNamesAndPaths)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

- toJSON()  
  Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.TemplateCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.TemplateCollectionData) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

- track()  
  Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

- untrack()  
  Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

## Property Details

### context

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value
- [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### items

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the loaded child items in this collection.

```typescript
readonly items: Word.Template[];
```

Property Value
- [Word.Template](/en-us/javascript/api/word/word.template)[]

## Method Details

### getCount()

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the number of items in the collection.

```typescript
getCount(): OfficeExtension.ClientResult<number>;
```

Returns
- [OfficeExtension.ClientResult](/en-us/javascript/api/office/officeextension.clientresult)<number>

Remarks
- [ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### getItemAt(index)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets a Template object by its index in the collection.

```typescript
getItemAt(index: number): Word.Template;
```

Parameters
- index  
  number

The index of the template to retrieve.

Returns
- [Word.Template](/en-us/javascript/api/word/word.template)

Remarks
- [ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### importBuildingBlocks()

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Imports the building blocks for all templates into Microsoft Word.

```typescript
importBuildingBlocks(): void;
```

Returns
- void

Remarks
- [ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### load(options)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(options?: Word.Interfaces.TemplateCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions): Word.TemplateCollection;
```

Parameters
- options  
  [Word.Interfaces.TemplateCollectionLoadOptions](/en-us/javascript/api/word/word.interfaces.templatecollectionloadoptions) & [Word.Interfaces.CollectionLoadOptions](/en-us/javascript/api/word/word.interfaces.collectionloadoptions)

Provides options for which properties of the object to load.

Returns
- [Word.TemplateCollection](/en-us/javascript/api/word/word.templatecollection)

### load(propertyNames)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.TemplateCollection;
```

Parameters
- propertyNames  
  string | string[]

A comma-delimited string or an array of strings that specify the properties to load.

Returns
- [Word.TemplateCollection](/en-us/javascript/api/word/word.templatecollection)

### load(propertyNamesAndPaths)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Word.TemplateCollection;
```

Parameters
- propertyNamesAndPaths  
  [OfficeExtension.LoadOption](/en-us/javascript/api/office/officeextension.loadoption)

propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

Returns
- [Word.TemplateCollection](/en-us/javascript/api/word/word.templatecollection)

### toJSON()

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.TemplateCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.TemplateCollectionData) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

```typescript
toJSON(): Word.Interfaces.TemplateCollectionData;
```

Returns
- [Word.Interfaces.TemplateCollectionData](/en-us/javascript/api/word/word.interfaces.templatecollectiondata)

### track()

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.TemplateCollection;
```

Returns
- [Word.TemplateCollection](/en-us/javascript/api/word/word.templatecollection)

### untrack()

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.TemplateCollection;
```

Returns
- [Word.TemplateCollection](/en-us/javascript/api/word/word.templatecollection)