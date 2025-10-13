# Word.BuildingBlockCategory class

Package: [word](/en-us/javascript/api/word)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents a category of building blocks in a Word document.

Extends: [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks

[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- buildingBlocks
  - Returns a BuildingBlockCollection object that represents the building blocks for the category.
- context
  - The request context associated with the object. This connects the add-in's process to the Office host application's process.
- index
  - Returns the position of the BuildingBlockCategory object in a collection.
- name
  - Returns the name of the BuildingBlockCategory object.
- type
  - Returns a BuildingBlockTypeItem object that represents the type of building block for the building block category.

## Methods

- load(options)
  - Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNames)
  - Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNamesAndPaths)
  - Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- toJSON()
  - Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.BuildingBlockCategory object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.BuildingBlockCategoryData) that contains shallow copies of any loaded child properties from the original object.
- track()
  - Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part o