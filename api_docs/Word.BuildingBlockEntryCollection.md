# Word.BuildingBlockEntryCollection class

Package: [word](/en-us/javascript/api/word)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents a collection of building block entries in a Word template.

- Extends: [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

## Properties
- [context](#context)
  - The request context associated with the object. This connects the add-in's process to the Office host application's process.

## Methods
- [add(name, type, category, range, description, insertType)](#addname-type-category-range-description-inserttype-1)  
  Creates a new building block entry in a template and returns a BuildingBlock object that represents the new building block entry.
- [add(name, type, category, range, description, insertType)](#addname-type-category-range-description-inserttype-2)  
  Creates a new building block entry in a template and returns a BuildingBlock object that represents the new building block entry.
- [getCount()](#getcount)  
  Returns the number of items in the collection.
- [getItemAt(index)](#getitematindex)  
  Returns a BuildingBlock object that represents the specified item in the collection.
- [load(propertyNames)](#loadpropertynames)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- [load(propertyNamesAndPaths)](#loadpropertynamesandpaths)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- [toJSON()](#tojson)  
  Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify().
- [track()](#track)  
  Track the object for automatic adjustment based on surrounding changes in the document.
- [untrack()](#untrack)  
  Release the memory associated with this object, if it has previously been tracked.

## Property Details

### context

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

- Property Value: [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

## Method Details

### add(name, type, category, range, description, insertType) (1)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Creates a new building block entry in a template and returns a BuildingBlock object that represents the new building block entry.

```typescript
add(name: string, type: Word.BuildingBlockType, category: string, range: Word.Range, description: string, insertType: Word.DocPartInsertType): Word.BuildingBlock;
```

- Parameters:
  - name: string  
    The name of the building block.
  - type: [Word.BuildingBlockType](/en-us/javascript/api/word/word.buildingblocktype)  
    The type of the building block.
  - category: string  
    The category of the building block.
  - range: [Word.Range](/en-us/javascript/api/word/word.range)  
    The range to insert the building block.
  - description: string  
    The description of the building block.
  - insertType: [Word.DocPartInsertType](/en-us/javascript/api/word/word.docpartinserttype)  
    How to insert the contents of the building block.

- Returns: [Word.BuildingBlock](/en-us/javascript/api/word/word.buildingblock)

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### add(name, type, category, range, description, insertType) (2)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Creates a new building block entry in a template and returns a BuildingBlock object that represents the new building block entry.

```typescript
add(name: string, type: "QuickParts" | "CoverPage" | "Equations" | "Footers" | "Headers" | "PageNumber" | "Tables" | "Watermarks" | "AutoText" | "TextBox" | "PageNumberTop" | "PageNumberBottom" | "PageNumberPage" | "TableOfContents" | "CustomQuickParts" | "CustomCoverPage" | "CustomEquations" | "CustomFooters" | "CustomHeaders" | "CustomPageNumber" | "CustomTables" | "CustomWatermarks" | "CustomAutoText" | "CustomTextBox" | "CustomPageNumberTop" | "CustomPageNumberBottom" | "CustomPageNumberPage" | "CustomTableOfContents" | "Custom1" | "Custom2" | "Custom3" | "Custom4" | "Custom5" | "Bibliography" | "CustomBibliography", category: string, range: Word.Range, description: string, insertType: "Content" | "Paragraph" | "Page"): Word.BuildingBlock;
```

- Parameters:
  - name: string  
    The name of the building block.
  - type: "QuickParts" | "CoverPage" | "Equations" | "Footers" | "Headers" | "PageNumber" | "Tables" | "Watermarks" | "AutoText" | "TextBox" | "PageNumberTop" | "PageNumberBottom" | "PageNumberPage" | "TableOfContents" | "CustomQuickParts" | "CustomCoverPage" | "CustomEquations" | "CustomFooters" | "CustomHeaders" | "CustomPageNumber" | "CustomTables" | "CustomWatermarks" | "CustomAutoText" | "CustomTextBox" | "CustomPageNumberTop" | "CustomPageNumberBottom" | "CustomPageNumberPage" | "CustomTableOfContents" | "Custom1" | "Custom2" | "Custom3" | "Custom4" | "Custom5" | "Bibliography" | "CustomBibliography"  
    The type of the building block.
  - category: string  
    The category of the building block.
  - range: [Word.Range](/en-us/javascript/api/word/word.range)  
    The range to insert the building block.
  - description: string  
    The description of the building block.
  - insertType: "Content" | "Paragraph" | "Page"  
    How to insert the contents of the building block.

- Returns: [Word.BuildingBlock](/en-us/javascript/api/word/word.buildingblock)

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### getCount()

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the number of items in the collection.

```typescript
getCount(): OfficeExtension.ClientResult<number>;
```

- Returns: [OfficeExtension.ClientResult](/en-us/javascript/api/office/officeextension.clientresult)<number>

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### getItemAt(index)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a BuildingBlock object that represents the specified item in the collection.

```typescript
getItemAt(index: number): Word.BuildingBlock;
```

- Parameters:
  - index: number  
    The index of the item to retrieve.

- Returns: [Word.BuildingBlock](/en-us/javascript/api/word/word.buildingblock)

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### load(propertyNames)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.BuildingBlockEntryCollection;
```

- Parameters:
  - propertyNames: string | string[]  
    A comma-delimited string or an array of strings that specify the properties to load.

- Returns: [Word.BuildingBlockEntryCollection](/en-us/javascript/api/word/word.buildingblockentrycollection)

### load(propertyNamesAndPaths)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
    select?: string;
    expand?: string;
}): Word.BuildingBlockEntryCollection;
```

- Parameters:
  - propertyNamesAndPaths: { select?: string; expand?: string; }  
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

- Returns: [Word.BuildingBlockEntryCollection](/en-us/javascript/api/word/word.buildingblockentrycollection)

### toJSON()

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.BuildingBlockEntryCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.BuildingBlockEntryCollectionData) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): {
    [key: string]: string;
};
```

- Returns: { [key: string]: string; }

### track()

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.BuildingBlockEntryCollection;
```

- Returns: [Word.BuildingBlockEntryCollection](/en-us/javascript/api/word/word.buildingblockentrycollection)

### untrack()

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.BuildingBlockEntryCollection;
```

- Returns: [Word.BuildingBlockEntryCollection](/en-us/javascript/api/word/word.buildingblockentrycollection)