# Word.Canvas class

Package: [word](/en-us/javascript/api/word)

Represents a canvas in the document. To get the corresponding Shape object, use Canvas.shape.

Extends
- [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks
[ API set: WordApiDesktop 1.2 ]

## Properties
- context  
  The request context associated with the object. This connects the add-in's process to the Office host application's process.

- id  
  Gets an integer that represents the canvas identifier.

- shape  
  Gets the Shape object associated with the canvas.

- shapes  
  Gets the collection of Shape objects. Currently, only text boxes, pictures, and geometric shapes are supported.

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
  Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.Canvas object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.CanvasData) that contains shallow copies of any loaded child properties from the original object.

- track()  
  Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

- untrack()  
  Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

## Property Details

### context
The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property value:
- [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### id
Gets an integer that represents the canvas identifier.

```typescript
readonly id: number;
```

Property value:
- number

Remarks
- [ API set: WordApiDesktop 1.2 ]

### shape
Gets the Shape object associated with the canvas.

```typescript
readonly shape: Word.Shape;
```

Property value:
- [Word.Shape](/en-us/javascript/api/word/word.shape)

Remarks
- [ API set: WordApiDesktop 1.2 ]

### shapes
Gets the collection of Shape objects. Currently, only text boxes, pictures, and geometric shapes are supported.

```typescript
readonly shapes: Word.ShapeCollection;
```

Property value:
- [Word.ShapeCollection](/en-us/javascript/api/word/word.shapecollection)

Remarks
- [ API set: WordApiDesktop 1.2 ]

## Method Details

### load(options)
Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(options?: Word.Interfaces.CanvasLoadOptions): Word.Canvas;
```

Parameters
- options: [Word.Interfaces.CanvasLoadOptions](/en-us/javascript/api/word/word.interfaces.canvasloadoptions)  
  Provides options for which properties of the object to load.

Returns
- [Word.Canvas](/en-us/javascript/api/word/word.canvas)

### load(propertyNames)
Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.Canvas;
```

Parameters
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns
- [Word.Canvas](/en-us/javascript/api/word/word.canvas)

### load(propertyNamesAndPaths)
Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.Canvas;
```

Parameters
- propertyNamesAndPaths:  
  {
  select?: string;
  expand?: string;
  }  
  propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

Returns
- [Word.Canvas](/en-us/javascript/api/word/word.canvas)

### set(properties, options)
Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.CanvasUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

Parameters
- properties: [Word.Interfaces.CanvasUpdateData](/en-us/javascript/api/word/word.interfaces.canvasupdatedata)  
  A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
- options: [OfficeExtension.UpdateOptions](/en-us/javascript/api/office/officeextension.updateoptions)  
  Provides an option to suppress errors if the properties object tries to set any read-only properties.

Returns
- void

### set(properties)
Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.Canvas): void;
```

Parameters
- properties: [Word.Canvas](/en-us/javascript/api/word/word.canvas)

Returns
- void

### toJSON()
Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.Canvas object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.CanvasData) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.CanvasData;
```

Returns
- [Word.Interfaces.CanvasData](/en-us/javascript/api/word/word.interfaces.canvasdata)

### track()
Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.Canvas;
```

Returns
- [Word.Canvas](/en-us/javascript/api/word/word.canvas)

### untrack()
Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.Canvas;
```

Returns
- [Word.Canvas](/en-us/javascript/api/word/word.canvas)