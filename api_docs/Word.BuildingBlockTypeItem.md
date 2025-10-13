# Word.BuildingBlockTypeItem class

Package: https://learn.microsoft.com/en-us/javascript/api/word

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents a type of building block in a Word document.

Extends: https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientobject

## Remarks

[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- categories  
  Returns a BuildingBlockCategoryCollection object that represents the categories for a building block type.

- context  
  The request context associated with the object. This connects the add-in's process to the Office host application's process.

- index  
  Returns the position of an item in a collection.

- name  
  Returns the localized name of a building block type.

## Methods

- load(options)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

- load(propertyNames)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

- load(propertyNamesAndPaths)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

- toJSON()  
  Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.BuildingBlockTypeItem object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.BuildingBlockTypeItemData) that contains shallow copies of any loaded child properties from the original object.

- track()  
  Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

- untrack()  
  Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

## Property Details

### categories

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a BuildingBlockCategoryCollection object that represents the categories for a building block type.

```typescript
readonly categories: Word.BuildingBlockCategoryCollection;
```

- Property Value: https://learn.microsoft.com/en-us/javascript/api/word/word.buildingblockcategorycollection

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### context

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

- Property Value: https://learn.microsoft.com/en-us/javascript/api/word/word.requestcontext

### index

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the position of an item in a collection.

```typescript
readonly index: number;
```

- Property Value: number

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### name

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the localized name of a building block type.

```typescript
readonly name: string;
```

- Property Value: string

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Method Details

### load(options)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(options?: Word.Interfaces.BuildingBlockTypeItemLoadOptions): Word.BuildingBlockTypeItem;
```

- Parameters:
  - options: https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.buildingblocktypeitemloadoptions  
    Provides options for which properties of the object to load.

- Returns: https://learn.microsoft.com/en-us/javascript/api/word/word.buildingblocktypeitem

### load(propertyNames)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.BuildingBlockTypeItem;
```

- Parameters:
  - propertyNames: string | string[]  
    A comma-delimited string or an array of strings that specify the properties to load.

- Returns: https://learn.microsoft.com/en-us/javascript/api/word/word.buildingblocktypeitem

### load(propertyNamesAndPaths)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
    select?: string;
    expand?: string;
}): Word.BuildingBlockTypeItem;
```

- Parameters:
  - propertyNamesAndPaths:  
    - select?: string  
    - expand?: string  
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

- Returns: https://learn.microsoft.com/en-us/javascript/api/word/word.buildingblocktypeitem

### toJSON()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.BuildingBlockTypeItem object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.BuildingBlockTypeItemData) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.BuildingBlockTypeItemData;
```

- Returns: https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.buildingblocktypeitemdata

### track()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.BuildingBlockTypeItem;
```

- Returns: https://learn.microsoft.com/en-us/javascript/api/word/word.buildingblocktypeitem

### untrack()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.BuildingBlockTypeItem;
```

- Returns: https://learn.microsoft.com/en-us/javascript/api/word/word.buildingblocktypeitem