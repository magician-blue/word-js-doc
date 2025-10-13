# Word.ShapeGroup class

Package: [word](/en-us/javascript/api/word)

Represents a shape group in the document. To get the corresponding Shape object, use ShapeGroup.shape.

Extends: [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks

[ API set: WordApiDesktop 1.2 ]

## Properties

- context  
  The request context associated with the object. This connects the add-in's process to the Office host application's process.

- id  
  Gets an integer that represents the shape group identifier.

- shape  
  Gets the Shape object associated with the group.

- shapes  
  Gets the collection of Shape objects. Currently, only text boxes, geometric shapes, and pictures are supported.

## Methods

- load(options)  
  Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

- load(propertyNames)  
  Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

- load(propertyNamesAndPaths)  
  Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

- set(properties, options)  
  Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

- set(properties)  
  Sets multiple properties on the object at the same time, based on an existing loaded object.

- toJSON()  
  Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.ShapeGroup` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.ShapeGroupData`) that contains shallow copies of any loaded child properties from the original object.

- track()  
  Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

- ungroup()  
  Ungroups any grouped shapes in the specified shape group.

- untrack()  
  Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

## Property Details

### context

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value  
[Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

---

### id

Gets an integer that represents the shape group identifier.

```typescript
readonly id: number;
```

Property Value  
number

Remarks  
[ API set: WordApiDesktop 1.2 ]

---

### shape

Gets the Shape object associated with the group.

```typescript
readonly shape: Word.Shape;
```

Property Value  
[Word.Shape](/en-us/javascript/api/word/word.shape)

Remarks  
[ API set: WordApiDesktop 1.2 ]

---

### shapes

Gets the collection of Shape objects. Currently, only text boxes, geometric shapes, and pictures are supported.

```typescript
readonly shapes: Word.ShapeCollection;
```

Property Value  
[Word.ShapeCollection](/en-us/javascript/api/word/word.shapecollection)

Remarks  
[ API set: WordApiDesktop 1.2 ]

## Method Details

### load(options)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(options?: Word.Interfaces.ShapeGroupLoadOptions): Word.ShapeGroup;
```

Parameters:
- options: [Word.Interfaces.ShapeGroupLoadOptions](/en-us/javascript/api/word/word.interfaces.shapegrouploadoptions)  
  Provides options for which properties of the object to load.

Returns  
[Word.ShapeGroup](/en-us/javascript/api/word/word.shapegroup)

---

### load(propertyNames)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.ShapeGroup;
```

Parameters:
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns  
[Word.ShapeGroup](/en-us/javascript/api/word/word.shapegroup)

---

### load(propertyNamesAndPaths)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.ShapeGroup;
```

Parameters:
- propertyNamesAndPaths:
  ```
  {
  select?: string;
  expand?: string;
  }
  ```
  `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

Returns  
[Word.ShapeGroup](/en-us/javascript/api/word/word.shapegroup)

---

### set(properties, options)

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.ShapeGroupUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

Parameters:
- properties: [Word.Interfaces.ShapeGroupUpdateData](/en-us/javascript/api/word/word.interfaces.shapegroupupdatedata)  
  A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
- options: [OfficeExtension.UpdateOptions](/en-us/javascript/api/office/officeextension.updateoptions)  
  Provides an option to suppress errors if the properties object tries to set any read-only properties.

Returns  
void

---

### set(properties)

Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.ShapeGroup): void;
```

Parameters:
- properties: [Word.ShapeGroup](/en-us/javascript/api/word/word.shapegroup)

Returns  
void

---

### toJSON()

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.ShapeGroup` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.ShapeGroupData`) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.ShapeGroupData;
```

Returns  
[Word.Interfaces.ShapeGroupData](/en-us/javascript/api/word/word.interfaces.shapegroupdata)

---

### track()

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.ShapeGroup;
```

Returns  
[Word.ShapeGroup](/en-us/javascript/api/word/word.shapegroup)

---

### ungroup()

Ungroups any grouped shapes in the specified shape group.

```typescript
ungroup(): Word.ShapeCollection;
```

Returns  
[Word.ShapeCollection](/en-us/javascript/api/word/word.shapecollection)

Remarks  
[ API set: WordApiDesktop 1.2 ]

---

### untrack()

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

```typescript
untrack(): Word.ShapeGroup;
```

Returns  
[Word.ShapeGroup](/en-us/javascript/api/word/word.shapegroup)