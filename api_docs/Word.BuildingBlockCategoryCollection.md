# Word.BuildingBlockCategoryCollection class

Package: [word](https://learn.microsoft.com/en-us/javascript/api/word)

Represents a collection of [Word.BuildingBlockCategory](https://learn.microsoft.com/en-us/javascript/api/word/word.buildingblockcategory) objects in a Word document.

Extends: [OfficeExtension.ClientObject](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientobject)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

## Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

## Properties
- context  
  The request context associated with the object. This connects the add-in's process to the Office host application's process.

## Methods
- getCount()  
  Returns the number of items in the collection.
- getItemAt(index)  
  Returns a BuildingBlockCategory object that represents the specified item in the collection.
- load(propertyNames)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNamesAndPaths)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- toJSON()  
  Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). Whereas the original Word.BuildingBlockCategoryCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.BuildingBlockCategoryCollectionData) that contains shallow copies of any loaded child properties from the original object.
- track()  
  Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
- untrack()  
  Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

## Property Details

### context
The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value: [Word.RequestContext](https://learn.microsoft.com/en-us/javascript/api/word/word.requestcontext)

## Method Details

### getCount()
Returns the number of items in the collection.

```typescript
getCount(): OfficeExtension.ClientResult<number>;
```

Returns: [OfficeExtension.ClientResult](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientresult)<number>

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

### getItemAt(index)
Returns a BuildingBlockCategory object that represents the specified item in the collection.

```typescript
getItemAt(index: number): Word.BuildingBlockCategory;
```

Parameters
- index: number  
  The index of the item to retrieve.

Returns: [Word.BuildingBlockCategory](https://learn.microsoft.com/en-us/javascript/api/word/word.buildingblockcategory)

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

### load(propertyNames)
Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.BuildingBlockCategoryCollection;
```

Parameters
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns: [Word.BuildingBlockCategoryCollection](https://learn.microsoft.com/en-us/javascript/api/word/word.buildingblockcategorycollection)

### load(propertyNamesAndPaths)
Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
    select?: string;
    expand?: string;
}): Word.BuildingBlockCategoryCollection;
```

Parameters
- propertyNamesAndPaths: { select?: string; expand?: string; }  
  propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

Returns: [Word.BuildingBlockCategoryCollection](https://learn.microsoft.com/en-us/javascript/api/word/word.buildingblockcategorycollection)

### toJSON()
Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.BuildingBlockCategoryCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.BuildingBlockCategoryCollectionData) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): {
    [key: string]: string;
};
```

Returns:
```
{
  [key: string]: string;
}
```

### track()
Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.BuildingBlockCategoryCollection;
```

Returns: [Word.BuildingBlockCategoryCollection](https://learn.microsoft.com/en-us/javascript/api/word/word.buildingblockcategorycollection)

### untrack()
Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.BuildingBlockCategoryCollection;
```

Returns: [Word.BuildingBlockCategoryCollection](https://learn.microsoft.com/en-us/javascript/api/word/word.buildingblockcategorycollection)