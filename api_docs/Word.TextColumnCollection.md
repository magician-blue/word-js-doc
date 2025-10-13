# Word.TextColumnCollection class

Package: [word](/en-us/javascript/api/word)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

A collection of [Word.TextColumn](/en-us/javascript/api/word/word.textcolumn) objects that represent all the columns of text in the document or a section of the document.

Extends
- [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties
- context — The request context associated with the object. This connects the add-in's process to the Office host application's process.
- items — Gets the loaded child items in this collection.

## Methods
- add(options) — Returns a TextColumn object that represents a new text column added to a section or document.
- getFlowDirection() — Gets the direction in which text flows from one text column to the next.
- getHasLineBetween() — Gets whether vertical lines appear between all the columns in the TextColumnCollection object.
- getIsEvenlySpaced() — Gets whether text columns are evenly spaced.
- getItem(index) — Gets a TextColumn by its index in the collection.
- load(options) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNames) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNamesAndPaths) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- setCount(numColumns) — Arranges text into the specified number of text columns.
- setFlowDirection(value) — Sets the direction in which text flows from one text column to the next.
- setFlowDirection(value) — Sets the direction in which text flows from one text column to the next.
- setHasLineBetween(value) — Sets whether vertical lines appear between all the columns in the TextColumnCollection object.
- setIsEvenlySpaced(value) — Sets whether text columns are evenly spaced.
- toJSON() — Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). Whereas the original Word.TextColumnCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.TextColumnCollectionData) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
- track() — Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
- untrack() — Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

## Property Details

### context
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value
- [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### items
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the loaded child items in this collection.

```typescript
readonly items: Word.TextColumn[];
```

Property Value
- [Word.TextColumn](/en-us/javascript/api/word/word.textcolumn)[]

## Method Details

### add(options)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a TextColumn object that represents a new text column added to a section or document.

```typescript
add(options?: Word.TextColumnAddOptions): Word.TextColumn;
```

Parameters
- options — [Word.TextColumnAddOptions](/en-us/javascript/api/word/word.textcolumnaddoptions)

Optional. Options for configuring the new text column.

Returns
- [Word.TextColumn](/en-us/javascript/api/word/word.textcolumn)

A TextColumn object that represents a new text column added to the document.

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### getFlowDirection()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the direction in which text flows from one text column to the next.

```typescript
getFlowDirection(): OfficeExtension.ClientResult<Word.FlowDirection>;
```

Returns
- [OfficeExtension.ClientResult](/en-us/javascript/api/office/officeextension.clientresult)<[Word.FlowDirection](/en-us/javascript/api/word/word.flowdirection)>

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### getHasLineBetween()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether vertical lines appear between all the columns in the TextColumnCollection object.

```typescript
getHasLineBetween(): OfficeExtension.ClientResult<boolean>;
```

Returns
- [OfficeExtension.ClientResult](/en-us/javascript/api/office/officeextension.clientresult)<boolean>

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### getIsEvenlySpaced()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether text columns are evenly spaced.

```typescript
getIsEvenlySpaced(): OfficeExtension.ClientResult<boolean>;
```

Returns
- [OfficeExtension.ClientResult](/en-us/javascript/api/office/officeextension.clientresult)<boolean>

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### getItem(index)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets a TextColumn by its index in the collection.

```typescript
getItem(index: number): Word.TextColumn;
```

Parameters
- index — number

A number that identifies the index location of a TextColumn object.

Returns
- [Word.TextColumn](/en-us/javascript/api/word/word.textcolumn)

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### load(options)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(options?: Word.Interfaces.TextColumnCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions): Word.TextColumnCollection;
```

Parameters
- options — [Word.Interfaces.TextColumnCollectionLoadOptions](/en-us/javascript/api/word/word.interfaces.textcolumncollectionloadoptions) & [Word.Interfaces.CollectionLoadOptions](/en-us/javascript/api/word/word.interfaces.collectionloadoptions)

Provides options for which properties of the object to load.

Returns
- [Word.TextColumnCollection](/en-us/javascript/api/word/word.textcolumncollection)

### load(propertyNames)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.TextColumnCollection;
```

Parameters
- propertyNames — string | string[]

A comma-delimited string or an array of strings that specify the properties to load.

Returns
- [Word.TextColumnCollection](/en-us/javascript/api/word/word.textcolumncollection)

### load(propertyNamesAndPaths)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Word.TextColumnCollection;
```

Parameters
- propertyNamesAndPaths — [OfficeExtension.LoadOption](/en-us/javascript/api/office/officeextension.loadoption)

propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

Returns
- [Word.TextColumnCollection](/en-us/javascript/api/word/word.textcolumncollection)

### setCount(numColumns)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Arranges text into the specified number of text columns.

```typescript
setCount(numColumns: number): void;
```

Parameters
- numColumns — number

The number of columns the text is to be arranged into.

Returns
- void

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### setFlowDirection(value)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets the direction in which text flows from one text column to the next.

```typescript
setFlowDirection(value: Word.FlowDirection): void;
```

Parameters
- value — [Word.FlowDirection](/en-us/javascript/api/word/word.flowdirection)

The flow direction to set.

Returns
- void

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### setFlowDirection(value)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets the direction in which text flows from one text column to the next.

```typescript
setFlowDirection(value: "LeftToRight" | "RightToLeft"): void;
```

Parameters
- value — "LeftToRight" | "RightToLeft"

The flow direction to set.

Returns
- void

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### setHasLineBetween(value)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets whether vertical lines appear between all the columns in the TextColumnCollection object.

```typescript
setHasLineBetween(value: boolean): void;
```

Parameters
- value — boolean

true to show vertical lines between columns.

Returns
- void

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### setIsEvenlySpaced(value)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets whether text columns are evenly spaced.

```typescript
setIsEvenlySpaced(value: boolean): void;
```

Parameters
- value — boolean

true to evenly space all the text columns in the document.

Returns
- void

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### toJSON()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.TextColumnCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.TextColumnCollectionData) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

```typescript
toJSON(): Word.Interfaces.TextColumnCollectionData;
```

Returns
- [Word.Interfaces.TextColumnCollectionData](/en-us/javascript/api/word/word.interfaces.textcolumncollectiondata)

### track()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.TextColumnCollection;
```

Returns
- [Word.TextColumnCollection](/en-us/javascript/api/word/word.textcolumncollection)

### untrack()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.TextColumnCollection;
```

Returns
- [Word.TextColumnCollection](/en-us/javascript/api/word/word.textcolumncollection)