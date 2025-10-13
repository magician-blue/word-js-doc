# Word.BuildingBlock class

Package: https://learn.microsoft.com/en-us/javascript/api/word

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents a building block in a template. A building block is pre-built content, similar to autotext, that may contain text, images, and formatting.

Extends
- OfficeExtension.ClientObject (https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientobject)

Remarks
- API set: WordApi BETA (PREVIEW ONLY) (https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- category  
  Returns a BuildingBlockCategory object that represents the category for the building block.

- context  
  The request context associated with the object. This connects the add-in's process to the Office host application's process.

- description  
  Specifies the description for the building block.

- id  
  Returns the internal identification number for the building block.

- index  
  Returns the position of this building block in a collection.

- insertType  
  Specifies a DocPartInsertType value that represents how to insert the contents of the building block into the document.

- name  
  Specifies the name of the building block.

- type  
  Returns a BuildingBlockTypeItem object that represents the type for the building block.

- value  
  Specifies the contents of the building block.

## Methods

- delete()  
  Deletes the building block.

- insert(range, richText)  
  Inserts the value of the building block into the document and returns a Range object that represents the contents of the building block within the document.

- load(options)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

- load(propertyNames)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

- load(propertyNamesAndPaths)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

- set(properties, options)  
  Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

- set(properties)  
  Sets multiple properties on the object at the same time, based on an existing loaded object.

- toJSON()  
  Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.BuildingBlock object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.BuildingBlockData) that contains shallow copies of any loaded child properties from the original object.

- track()  
  Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

- untrack()  
  Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

## Property Details

### category
Returns a BuildingBlockCategory object that represents the category for the building block.

```typescript
readonly category: Word.BuildingBlockCategory;
```

Property value
- Word.BuildingBlockCategory (https://learn.microsoft.com/en-us/javascript/api/word/word.buildingblockcategory)

Remarks
- API set: WordApi BETA (PREVIEW ONLY) (https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### context
The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property value
- Word.RequestContext (https://learn.microsoft.com/en-us/javascript/api/word/word.requestcontext)

### description
Specifies the description for the building block.

```typescript
description: string;
```

Property value
- string

Remarks
- API set: WordApi BETA (PREVIEW ONLY) (https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### id
Returns the internal identification number for the building block.

```typescript
readonly id: string;
```

Property value
- string

Remarks
- API set: WordApi BETA (PREVIEW ONLY) (https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### index
Returns the position of this building block in a collection.

```typescript
readonly index: number;
```

Property value
- number

Remarks
- API set: WordApi BETA (PREVIEW ONLY) (https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### insertType
Specifies a DocPartInsertType value that represents how to insert the contents of the building block into the document.

```typescript
insertType: Word.DocPartInsertType | "Content" | "Paragraph" | "Page";
```

Property value
- Word.DocPartInsertType (https://learn.microsoft.com/en-us/javascript/api/word/word.docpartinserttype) | "Content" | "Paragraph" | "Page"

Remarks
- API set: WordApi BETA (PREVIEW ONLY) (https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### name
Specifies the name of the building block.

```typescript
name: string;
```

Property value
- string

Remarks
- API set: WordApi BETA (PREVIEW ONLY) (https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### type
Returns a BuildingBlockTypeItem object that represents the type for the building block.

```typescript
readonly type: Word.BuildingBlockTypeItem;
```

Property value
- Word.BuildingBlockTypeItem (https://learn.microsoft.com/en-us/javascript/api/word/word.buildingblocktypeitem)

Remarks
- API set: WordApi BETA (PREVIEW ONLY) (https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### value
Specifies the contents of the building block.

```typescript
value: string;
```

Property value
- string

Remarks
- API set: WordApi BETA (PREVIEW ONLY) (https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Method Details

### delete()
Deletes the building block.

```typescript
delete(): void;
```

Returns
- void

Remarks
- API set: WordApi BETA (PREVIEW ONLY) (https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### insert(range, richText)
Inserts the value of the building block into the document and returns a Range object that represents the contents of the building block within the document.

```typescript
insert(range: Word.Range, richText: boolean): Word.Range;
```

Parameters
- range: Word.Range (https://learn.microsoft.com/en-us/javascript/api/word/word.range)  
  The range where the building block should be inserted.
- richText: boolean  
  Indicates whether to insert as rich text.

Returns
- Word.Range (https://learn.microsoft.com/en-us/javascript/api/word/word.range)

Remarks
- API set: WordApi BETA (PREVIEW ONLY) (https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### load(options)
Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(options?: Word.Interfaces.BuildingBlockLoadOptions): Word.BuildingBlock;
```

Parameters
- options: Word.Interfaces.BuildingBlockLoadOptions (https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.buildingblockloadoptions)  
  Provides options for which properties of the object to load.

Returns
- Word.BuildingBlock (https://learn.microsoft.com/en-us/javascript/api/word/word.buildingblock)

### load(propertyNames)
Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.BuildingBlock;
```

Parameters
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns
- Word.BuildingBlock (https://learn.microsoft.com/en-us/javascript/api/word/word.buildingblock)

### load(propertyNamesAndPaths)
Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.BuildingBlock;
```

Parameters
- propertyNamesAndPaths: { select?: string; expand?: string; }  
  propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

Returns
- Word.BuildingBlock (https://learn.microsoft.com/en-us/javascript/api/word/word.buildingblock)

### set(properties, options)
Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.BuildingBlockUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

Parameters
- properties: Word.Interfaces.BuildingBlockUpdateData (https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.buildingblockupdatedata)  
  A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
- options: OfficeExtension.UpdateOptions (https://learn.microsoft.com/en-us/javascript/api/office/officeextension.updateoptions)  
  Provides an option to suppress errors if the properties object tries to set any read-only properties.

Returns
- void

### set(properties)
Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.BuildingBlock): void;
```

Parameters
- properties: Word.BuildingBlock (https://learn.microsoft.com/en-us/javascript/api/word/word.buildingblock)

Returns
- void

### toJSON()
Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.BuildingBlock object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.BuildingBlockData) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.BuildingBlockData;
```

Returns
- Word.Interfaces.BuildingBlockData (https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.buildingblockdata)

### track()
Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.BuildingBlock;
```

Returns
- Word.BuildingBlock (https://learn.microsoft.com/en-us/javascript/api/word/word.buildingblock)

### untrack()
Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.BuildingBlock;
```

Returns
- Word.BuildingBlock (https://learn.microsoft.com/en-us/javascript/api/word/word.buildingblock)