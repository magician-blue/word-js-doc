# Word.BuildingBlockCollection class

Package: [word](/en-us/javascript/api/word)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents a collection of [Word.BuildingBlock](/en-us/javascript/api/word/word.buildingblock) objects for a specific building block type and category in a template.

Extends: [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- context
  - The request context associated with the object. This connects the add-in's process to the Office host application's process.

## Methods

- add(name, range, description, insertType)
  - Creates a new building block and returns a BuildingBlock object.
- add(name, range, description, insertType)
  - Creates a new building block and returns a BuildingBlock object.
- getCount()
  - Returns the number of items in the collection.
- getItemAt(index)
  - Returns a BuildingBlock object that represents the specified item in the collection.
- load(propertyNames)
  - Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNamesAndPaths)
  - Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- toJSON()
  - Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.BuildingBlockCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.BuildingBlockCollectionData) that contains shallow copies of any loaded child properties from the original object.
- track()
  - Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
- untrack()
  - Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

## Property Details

### context

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

- Property Value: [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

## Method Details

### add(name, range, description, insertType)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Creates a new building block and returns a BuildingBlock object.

```typescript
add(name: string, range: Word.Range, description: string, insertType: Word.DocPartInsertType): Word.BuildingBlock;
```

- Parameters:
  - name: string
    - The name of the building block.
  - range: [Word.Range](/en-us/javascript/api/word/word.range)
    - The range to insert the building block.
  - description: string
    - The description of the building block.
  - insertType: [Word.DocPartInsertType](/en-us/javascript/api/word/word.docpartinserttype)
    - How to insert the contents of the building block.
- Returns: [Word.BuildingBlock](/en-us/javascript/api/word/word.buildingblock)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### add(name, range, description, insertType)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Creates a new building block and returns a BuildingBlock object.

```typescript
add(name: string, range: Word.Range, description: string, insertType: "Content" | "Paragraph" | "Page"): Word.BuildingBlock;
```

- Parameters:
  - name: string
    - The name of the building block.
  - range: [Word.Range](/en-us/javascript/api/word/word.range)
    - The range to insert the building block.
  - description: string
    - The description of the building block.
  - insertType: "Content" | "Paragraph" | "Page"
    - How to insert the contents of the building block.
- Returns: [Word.BuildingBlock](/en-us/javascript/api/word/word.buildingblock)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### getCount()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the number of items in the collection.

```typescript
getCount(): OfficeExtension.ClientResult<number>;
```

- Returns: [OfficeExtension.ClientResult](/en-us/javascript/api/office/officeextension.clientresult)<number>

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### getItemAt(index)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a BuildingBlock object that represents the specified item in the collection.

```typescript
getItemAt(index: number): Word.BuildingBlock;
```

- Parameters:
  - index: number
    - The index of the item to retrieve.
- Returns: [Word.BuildingBlock](/en-us/javascript/api/word/word.buildingblock)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### load(propertyNames)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.BuildingBlockCollection;
```

- Parameters:
  - propertyNames: string | string[]
    - A comma-delimited string or an array of strings that specify the properties to load.
- Returns: [Word.BuildingBlockCollection](/en-us/javascript/api/word/word.buildingblockcollection)

### load(propertyNamesAndPaths)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
    select?: string;
    expand?: string;
}): Word.BuildingBlockCollection;
```

- Parameters:
  - propertyNamesAndPaths: { select?: string; expand?: string; }
    - propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.
- Returns: [Word.BuildingBlockCollection](/en-us/javascript/api/word/word.buildingblockcollection)

### toJSON()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.BuildingBlockCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.BuildingBlockCollectionData) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): {
    [key: string]: string;
};
```

- Returns: { [key: string]: string; }

### track()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.BuildingBlockCollection;
```

- Returns: [Word.BuildingBlockCollection](/en-us/javascript/api/word/word.buildingblockcollection)

### untrack()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.BuildingBlockCollection;
```

- Returns: [Word.BuildingBlockCollection](/en-us/javascript/api/word/word.buildingblockcollection)