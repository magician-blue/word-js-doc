# Word.BuildingBlockTypeItemCollection class

- Package: [word](/en-us/javascript/api/word)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents a collection of building block types in a Word document.

- Extends: [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks

[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- context  
  The request context associated with the object. This connects the add-in's process to the Office host application's process.

## Methods

- getByType(type)  
  Gets a BuildingBlockTypeItem object by its type in the collection.

- getByType(type)  
  Gets a BuildingBlockTypeItem object by its type in the collection.

- getCount()  
  Returns the number of items in the collection.

- load(propertyNames)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

- load(propertyNamesAndPaths)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

- toJSON()  
  Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.BuildingBlockTypeItemCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.BuildingBlockTypeItemCollectionData) that contains shallow copies of any loaded child properties from the original object.

- track()  
  Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

- untrack()  
  Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

## Property Details

### context

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value: [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

## Method Details

### getByType(type)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets a BuildingBlockTypeItem object by its type in the collection.

```typescript
getByType(type: Word.BuildingBlockType): Word.BuildingBlockTypeItem;
```

Parameters:
- type: [Word.BuildingBlockType](/en-us/javascript/api/word/word.buildingblocktype)  
  Specifies the building block type of the item in the collection.

Returns: [Word.BuildingBlockTypeItem](/en-us/javascript/api/word/word.buildingblocktypeitem)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### getByType(type)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets a BuildingBlockTypeItem object by its type in the collection.

```typescript
getByType(type: "QuickParts" | "CoverPage" | "Equations" | "Footers" | "Headers" | "PageNumber" | "Tables" | "Watermarks" | "AutoText" | "TextBox" | "PageNumberTop" | "PageNumberBottom" | "PageNumberPage" | "TableOfContents" | "CustomQuickParts" | "CustomCoverPage" | "CustomEquations" | "CustomFooters" | "CustomHeaders" | "CustomPageNumber" | "CustomTables" | "CustomWatermarks" | "CustomAutoText" | "CustomTextBox" | "CustomPageNumberTop" | "CustomPageNumberBottom" | "CustomPageNumberPage" | "CustomTableOfContents" | "Custom1" | "Custom2" | "Custom3" | "Custom4" | "Custom5" | "Bibliography" | "CustomBibliography"): Word.BuildingBlockTypeItem;
```

Parameters:
- type: "QuickParts" | "CoverPage" | "Equations" | "Footers" | "Headers" | "PageNumber" | "Tables" | "Watermarks" | "AutoText" | "TextBox" | "PageNumberTop" | "PageNumberBottom" | "PageNumberPage" | "TableOfContents" | "CustomQuickParts" | "CustomCoverPage" | "CustomEquations" | "CustomFooters" | "CustomHeaders" | "CustomPageNumber" | "CustomTables" | "CustomWatermarks" | "CustomAutoText" | "CustomTextBox" | "CustomPageNumberTop" | "CustomPageNumberBottom" | "CustomPageNumberPage" | "CustomTableOfContents" | "Custom1" | "Custom2" | "Custom3" | "Custom4" | "Custom5" | "Bibliography" | "CustomBibliography"  
  Specifies the building block type of the item in the collection.

Returns: [Word.BuildingBlockTypeItem](/en-us/javascript/api/word/word.buildingblocktypeitem)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### getCount()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the number of items in the collection.

```typescript
getCount(): OfficeExtension.ClientResult<number>;
```

Returns: [OfficeExtension.ClientResult](/en-us/javascript/api/office/officeextension.clientresult)<number>

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### load(propertyNames)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.BuildingBlockTypeItemCollection;
```

Parameters:
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns: [Word.BuildingBlockTypeItemCollection](/en-us/javascript/api/word/word.buildingblocktypeitemcollection)

### load(propertyNamesAndPaths)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.BuildingBlockTypeItemCollection;
```

Parameters:
- propertyNamesAndPaths:  
  {
  select?: string;
  expand?: string;
  }  
  propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

Returns: [Word.BuildingBlockTypeItemCollection](/en-us/javascript/api/word/word.buildingblocktypeitemcollection)

### toJSON()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.BuildingBlockTypeItemCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.BuildingBlockTypeItemCollectionData) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): {
            [key: string]: string;
        };
```

Returns:  
{
[key: string]: string;
}

### track()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.BuildingBlockTypeItemCollection;
```

Returns: [Word.BuildingBlockTypeItemCollection](/en-us/javascript/api/word/word.buildingblocktypeitemcollection)

### untrack()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.BuildingBlockTypeItemCollection;
```

Returns: [Word.BuildingBlockTypeItemCollection](/en-us/javascript/api/word/word.buildingblocktypeitemcollection)