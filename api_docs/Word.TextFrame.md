# Word.TextFrame class

Represents the text frame of a shape object.

- Package: [word](https://learn.microsoft.com/en-us/javascript/api/word)
- Extends: [OfficeExtension.ClientObject](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientobject)

## Remarks

[API set: WordApiDesktop 1.2](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- autoSizeSetting  
  The automatic sizing settings for the text frame. A text frame can be set to automatically fit the text to the text frame, to automatically fit the text frame to the text, or not perform any automatic sizing.
- bottomMargin  
  Represents the bottom margin, in points, of the text frame.
- context  
  The request context associated with the object. This connects the add-in's process to the Office host application's process.
- hasText  
  Specifies if the text frame contains text.
- leftMargin  
  Represents the left margin, in points, of the text frame.
- noTextRotation  
  Returns True if text in the text frame shouldn't rotate when the shape is rotated.
- orientation  
  Represents the angle to which the text is oriented for the text frame. See Word.ShapeTextOrientation for details.
- rightMargin  
  Represents the right margin, in points, of the text frame.
- topMargin  
  Represents the top margin, in points, of the text frame.
- verticalAlignment  
  Represents the vertical alignment of the text frame. See Word.ShapeTextVerticalAlignment for details.
- wordWrap  
  Determines whether lines break automatically to fit text inside the shape.

## Methods

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
  Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). JSON.stringify, in turn, calls the toJSON method of the object that's passed to it. Whereas the original Word.TextFrame object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.TextFrameData) that contains shallow copies of any loaded child properties from the original object.
- track()  
  Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
- untrack()  
  Release the memory associated with this object, if it has previously been tracked. This call is shorthand for contex