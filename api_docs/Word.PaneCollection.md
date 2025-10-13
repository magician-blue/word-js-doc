# Word.PaneCollection class

Package: [word](/en-us/javascript/api/word)

Represents the collection of pane.

Extends [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks

[ API set: WordApiDesktop 1.2 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- [context](#word-word-panecollection-context-member)  
  The request context associated with the object. This connects the add-in's process to the Office host application's process.

- [items](#word-word-panecollection-items-member)  
  Gets the loaded child items in this collection.

## Methods

- [getFirst()](#word-word-panecollection-getfirst-member1)  
  Gets the first pane in this collection. Throws an `ItemNotFound` error if this collection is empty.

- [getFirstOrNullObject()](#word-word-panecollection-getfirstornullobject-member1)  
  Gets the first pane in this collection. If this collection is empty, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

- [load(propertyNames)](#word-word-panecollection-load-member1)  
  Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

- [load(propertyNamesAndPaths)](#word-word-panecollection-load-member2)  
  Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

- [toJSON()](#word-word-panecollection-tojson-member1)  
  Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.PaneCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.PaneCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

- [track()](#word-word-panecollection-track-member1)  
  Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

- [untrack()](#word-word-panecollection-untrack-member1)  
  Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

## Property Details

### <a id="word-word-panecollection-context-member"></a>context

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value  
[Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### <a id="word-word-panecollection-items-member"></a>items

Gets the loaded child items in this collection.

```typescript
readonly items: Word.Pane[];
```

Property Value  
[Word.Pane](/en-us/javascript/api/word/word.pane)[]

## Method Details

### <a id="word-word-panecollection-getfirst-member(1)"></a>getFirst()

Gets the first pane in this collection. Throws an `ItemNotFound` error if this collection is empty.

```typescript
getFirst(): Word.Pane;
```

Returns  
[Word.Pane](/en-us/javascript/api/word/word.pane)

Remarks  
[ API set: WordApiDesktop 1.2 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### <a id="word-word-panecollection-getfirstornullobject-member(1)"></a>getFirstOrNullObject()

Gets the first pane in this collection. If this collection is empty, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
getFirstOrNullObject(): Word.Pane;
```

Returns  
[Word.Pane](/en-us/javascript/api/word/word.pane)

Remarks  
[ API set: WordApiDesktop 1.2 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### <a id="word-word-panecollection-load-member(1)"></a>load(propertyNames)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.PaneCollection;
```

Parameters  
- propertyNames  
  string | string[]

A comma-delimited string or an array of strings that specify the properties to load.

Returns  
[Word.PaneCollection](/en-us/javascript/api/word/word.panecollection)

### <a id="word-word-panecollection-load-member(2)"></a>load(propertyNamesAndPaths)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Word.PaneCollection;
```

Parameters  
- propertyNamesAndPaths  
  [OfficeExtension.LoadOption](/en-us/javascript/api/office/officeextension.loadoption)

`propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

Returns  
[Word.PaneCollection](/en-us/javascript/api/word/word.panecollection)

### <a id="word-word-panecollection-tojson-member(1)"></a>toJSON()

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.PaneCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.PaneCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

```typescript
toJSON(): Word.Interfaces.PaneCollectionData;
```

Returns  
[Word.Interfaces.PaneCollectionData](/en-us/javascript/api/word/word.interfaces.panecollectiondata)

### <a id="word-word-panecollection-track-member(1)"></a>track()

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.PaneCollection;
```

Returns  
[Word.PaneCollection](/en-us/javascript/api/word/word.panecollection)

### <a id="word-word-panecollection-untrack-member(1)"></a>untrack()

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

```typescript
untrack(): Word.PaneCollection;
```

Returns  
[Word.PaneCollection](/en-us/javascript/api/word/word.panecollection)