# Word.TableColumnCollection class

Package: [word](/en-us/javascript/api/word)

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents a collection of [Word.TableColumn](/en-us/javascript/api/word/word.tablecolumn) objects in a Word document.

Extends: [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks

[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- [context](#context)  
  The request context associated with the object. This connects the add-in's process to the Office host application's process.
- [items](#items)  
  Gets the loaded child items in this collection.

## Methods

- [add(beforeColumn)](#addbeforecolumn)  
  Returns a TableColumn object that represents a column added to a table.
- [autoFit()](#autofit)  
  Changes the width of a table column to accommodate the width of the text without changing the way text wraps in the cells.
- [delete()](#delete)  
  Deletes the specified columns.
- [distributeWidth()](#distributewidth)  
  Adjusts the width of the specified columns so that they are equal.
- [load(options)](#loadoptions)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- [load(propertyNames)](#loadpropertynames)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- [load(propertyNamesAndPaths)](#loadpropertynamesandpaths)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- [select()](#select)  
  Selects the specified table columns.
- [setWidth(columnWidth, rulerStyle)](#setwidthcolumnwidth-rulerstyle)  
  Sets the width of columns in a table.
- [setWidth(columnWidth, rulerStyle)](#setwidthcolumnwidth-rulerstyle-1)  
  Sets the width of columns in a table.
- [toJSON()](#tojson)  
  Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify().
- [track()](#track)  
  Track the object for automatic adjustment based on surrounding changes in the document.
- [untrack()](#untrack)  
  Release the memory associated with this object, if it has previously been tracked.

## Property Details

### context

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value: [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### items

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the loaded child items in this collection.

```typescript
readonly items: Word.TableColumn[];
```

Property Value: [Word.TableColumn](/en-us/javascript/api/word/word.tablecolumn)[]

## Method Details

### add(beforeColumn)

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a TableColumn object that represents a column added to a table.

```typescript
add(beforeColumn?: Word.TableColumn): Word.TableColumn;
```

Parameters:
- beforeColumn: [Word.TableColumn](/en-us/javascript/api/word/word.tablecolumn)  
  Optional. The column before which the new column is added.

Returns: [Word.TableColumn](/en-us/javascript/api/word/word.tablecolumn)  
A new TableColumn object.

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### autoFit()

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Changes the width of a table column to accommodate the width of the text without changing the way text wraps in the cells.

```typescript
autoFit(): void;
```

Returns: void

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### delete()

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Deletes the specified columns.

```typescript
delete(): void;
```

Returns: void

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### distributeWidth()

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Adjusts the width of the specified columns so that they are equal.

```typescript
distributeWidth(): void;
```

Returns: void

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### load(options)

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(options?: Word.Interfaces.TableColumnCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions): Word.TableColumnCollection;
```

Parameters:
- options: [Word.Interfaces.TableColumnCollectionLoadOptions](/en-us/javascript/api/word/word.interfaces.tablecolumncollectionloadoptions) & [Word.Interfaces.CollectionLoadOptions](/en-us/javascript/api/word/word.interfaces.collectionloadoptions)  
  Provides options for which properties of the object to load.

Returns: [Word.TableColumnCollection](/en-us/javascript/api/word/word.tablecolumncollection)

### load(propertyNames)

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.TableColumnCollection;
```

Parameters:
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns: [Word.TableColumnCollection](/en-us/javascript/api/word/word.tablecolumncollection)

### load(propertyNamesAndPaths)

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Word.TableColumnCollection;
```

Parameters:
- propertyNamesAndPaths: [OfficeExtension.LoadOption](/en-us/javascript/api/office/officeextension.loadoption)  
  `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

Returns: [Word.TableColumnCollection](/en-us/javascript/api/word/word.tablecolumncollection)

### select()

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Selects the specified table columns.

```typescript
select(): void;
```

Returns: void

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### setWidth(columnWidth, rulerStyle)

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets the width of columns in a table.

```typescript
setWidth(columnWidth: number, rulerStyle: Word.RulerStyle): void;
```

Parameters:
- columnWidth: number  
  The width to set.
- rulerStyle: [Word.RulerStyle](/en-us/javascript/api/word/word.rulerstyle)  
  The ruler style to apply.

Returns: void

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### setWidth(columnWidth, rulerStyle)

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets the width of columns in a table.

```typescript
setWidth(columnWidth: number, rulerStyle: "None" | "Proportional" | "FirstColumn" | "SameWidth"): void;
```

Parameters:
- columnWidth: number  
  The width to set.
- rulerStyle: "None" | "Proportional" | "FirstColumn" | "SameWidth"  
  The ruler style to apply.

Returns: void

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### toJSON()

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.TableColumnCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.TableColumnCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

```typescript
toJSON(): Word.Interfaces.TableColumnCollectionData;
```

Returns: [Word.Interfaces.TableColumnCollectionData](/en-us/javascript/api/word/word.interfaces.tablecolumncollectiondata)

### track()

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.TableColumnCollection;
```

Returns: [Word.TableColumnCollection](/en-us/javascript/api/word/word.tablecolumncollection)

### untrack()

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

```typescript
untrack(): Word.TableColumnCollection;
```

Returns: [Word.TableColumnCollection](/en-us/javascript/api/word/word.tablecolumncollection)