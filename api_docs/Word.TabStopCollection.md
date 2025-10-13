# Word.TabStopCollection class

- Package: [word](/en-us/javascript/api/word)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents a collection of [tab stops](/en-us/javascript/api/word/word.tabstop) in a Word document.

- Extends: [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

## Properties
- context: The request context associated with the object. This connects the add-in's process to the Office host application's process.
- items: Gets the loaded child items in this collection.

## Methods
- add(position, options): Returns a TabStop object that represents a custom tab stop added to the paragraph.
- after(Position): Returns the next TabStop object to the right of the specified position.
- before(Position): Returns the next TabStop object to the left of the specified position.
- clearAll(): Clears all the custom tab stops from the paragraph.
- getItem(index): Gets a TabStop object by its index in the collection.
- load(options): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNames): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNamesAndPaths): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- toJSON(): Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.TabStopCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.TabStopCollectionData) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
- track(): Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
- untrack(): Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

## Property Details

### context
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value: [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### items
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the loaded child items in this collection.

```typescript
readonly items: Word.TabStop[];
```

Property Value: [Word.TabStop](/en-us/javascript/api/word/word.tabstop)[]

## Method Details

### add(position, options)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `TabStop` object that represents a custom tab stop added to the paragraph.

```typescript
add(position: number, options?: Word.TabStopAddOptions): Word.TabStop;
```

Parameters:
- position: number  
  - The position of the tab stop.
- options: [Word.TabStopAddOptions](/en-us/javascript/api/word/word.tabstopaddoptions)  
  - Optional. The options to further configure the new tab stop.

Returns: [Word.TabStop](/en-us/javascript/api/word/word.tabstop)

Remarks:
[ API set: WordApi BETA (PREVIEW ONLY) ]

### after(Position)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the next `TabStop` object to the right of the specified position.

```typescript
after(Position: number): Word.TabStop;
```

Parameters:
- Position: number  
  - The position to check.

Returns: [Word.TabStop](/en-us/javascript/api/word/word.tabstop)

Remarks:
[ API set: WordApi BETA (PREVIEW ONLY) ]

### before(Position)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the next `TabStop` object to the left of the specified position.

```typescript
before(Position: number): Word.TabStop;
```

Parameters:
- Position: number  
  - The position to check.

Returns: [Word.TabStop](/en-us/javascript/api/word/word.tabstop)

Remarks:
[ API set: WordApi BETA (PREVIEW ONLY) ]

### clearAll()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Clears all the custom tab stops from the paragraph.

```typescript
clearAll(): void;
```

Returns: void

Remarks:
[ API set: WordApi BETA (PREVIEW ONLY) ]

### getItem(index)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets a `TabStop` object by its index in the collection.

```typescript
getItem(index: number): Word.TabStop;
```

Parameters:
- index: number  
  - A number that identifies the index location of a `TabStop` object.

Returns: [Word.TabStop](/en-us/javascript/api/word/word.tabstop)

Remarks:
[ API set: WordApi BETA (PREVIEW ONLY) ]

### load(options)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(options?: Word.Interfaces.TabStopCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions): Word.TabStopCollection;
```

Parameters:
- options: [Word.Interfaces.TabStopCollectionLoadOptions](/en-us/javascript/api/word/word.interfaces.tabstopcollectionloadoptions) & [Word.Interfaces.CollectionLoadOptions](/en-us/javascript/api/word/word.interfaces.collectionloadoptions)  
  - Provides options for which properties of the object to load.

Returns: [Word.TabStopCollection](/en-us/javascript/api/word/word.tabstopcollection)

### load(propertyNames)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.TabStopCollection;
```

Parameters:
- propertyNames: string | string[]  
  - A comma-delimited string or an array of strings that specify the properties to load.

Returns: [Word.TabStopCollection](/en-us/javascript/api/word/word.tabstopcollection)

### load(propertyNamesAndPaths)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Word.TabStopCollection;
```

Parameters:
- propertyNamesAndPaths: [OfficeExtension.LoadOption](/en-us/javascript/api/office/officeextension.loadoption)  
  - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

Returns: [Word.TabStopCollection](/en-us/javascript/api/word/word.tabstopcollection)

### toJSON()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.TabStopCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.TabStopCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

```typescript
toJSON(): Word.Interfaces.TabStopCollectionData;
```

Returns: [Word.Interfaces.TabStopCollectionData](/en-us/javascript/api/word/word.interfaces.tabstopcollectiondata)

### track()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.TabStopCollection;
```

Returns: [Word.TabStopCollection](/en-us/javascript/api/word/word.tabstopcollection)

### untrack()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

```typescript
untrack(): Word.TabStopCollection;
```

Returns: [Word.TabStopCollection](/en-us/javascript/api/word/word.tabstopcollection)