# Word.ReviewerCollection class

Package: https://learn.microsoft.com/en-us/javascript/api/word

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

A collection of Word.Reviewer objects that represents the reviewers of one or more documents. The ReviewerCollection object contains the names of all reviewers who have reviewed documents opened or edited on a computer.

- Extends: https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientobject

## Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

## Properties

- context
  - The request context associated with the object. This connects the add-in's process to the Office host application's process.
- items
  - Gets the loaded child items in this collection.

## Methods

- getItem(index)
  - Returns a Reviewer object that represents the specified item in the collection.
- load(options)
  - Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNames)
  - Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNamesAndPaths)
  - Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- toJSON()
  - Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.ReviewerCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.ReviewerCollectionData) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
- track()
  - Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
- untrack()
  - Release the memory associated with this 