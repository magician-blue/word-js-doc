# Word.ShapeCollection class

- Package: [word](/en-us/javascript/api/word)

Contains a collection of [Word.Shape](/en-us/javascript/api/word/word.shape) objects. Currently, only the following shapes are supported: text boxes, geometric shapes, groups, pictures, and canvases.

- Extends: [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks

[ API set: WordApiDesktop 1.2 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples

```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-shapes-text-boxes.yaml

await Word.run(async (context) => {
  // Gets text boxes in main document.
  const shapes: Word.ShapeCollection = context.document.body.shapes;
  shapes.load();
  await context.sync();

  if (shapes.items.length > 0) {
    shapes.items.forEach(function(shape, index) {
      if (shape.type === Word.ShapeType.textBox) {
        console.log(`Shape ${index} in the main document has a text box. Properties:`, shape);
      }
    });
  } else {
    console.log("No shapes found in main document.");
  }
});
```

## Properties

- context  
  The request context associated with the object. This connects the add-in's process to the Office host application's process.

- items  
  Gets the loaded child items in this collection.

## Methods

- getByGeometricTypes(types)  
  Gets the shapes that have the specified geometric types. Only applied to geometric shapes.

- getById(id)  
  Gets a shape by its identifier. Throws an `ItemNotFound` error if there isn't a shape with the identifier in this collection.

- getByIdOrNullObject(id)  
  Gets a shape by its identifier. If there isn't a shape with the identifier in this collection, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

- getByIds(ids)  
  Gets the shapes by the identifiers.

- getByNames(names)  
  Gets the shapes that have the specified names.

- getByTypes(types)  
  Gets the shapes that have the specified types.

- getFirst()  
  Gets the first shape in this collection. Throws an `ItemNotFound` error if this collection is empty.

- getFirstOrNullObject()  
  Gets the first shape in this collection. If this collection is empty, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

- group()  
  Groups floating shapes in this collection, inline shapes will be skipped. Returns a Shape object that represents the new group of shapes.

- load(options)  
  Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

- load(propertyNames)  
  Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

- load(propertyNamesAndPaths)  
  Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

- toJSON()  
  Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.ShapeCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.ShapeCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

- track()  
  Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

- untrack()  
  Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

## Property Details

### context

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property value: [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### items

Gets the loaded child items in this collection.

```typescript
readonly items: Word.Shape[];
```

Property value: [Word.Shape](/en-us/javascript/api/word/word.shape)[]

## Method Details

### getByGeometricTypes(types)

Gets the shapes that have the specified geometric types. Only applied to geometric shapes.

```typescript
getByGeometricTypes(types: Word.GeometricShapeType[]): Word.ShapeCollection;
```

Parameters:
- types: [Word.GeometricShapeType](/en-us/javascript/api/word/word.geometricshapetype)[]  
  Required. An array of geometric shape subtypes.

Returns: [Word.ShapeCollection](/en-us/javascript/api/word/word.shapecollection)

Remarks: [ API set: WordApiDesktop 1.2 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### getById(id)

Gets a shape by its identifier. Throws an `ItemNotFound` error if there isn't a shape with the identifier in this collection.

```typescript
getById(id: number): Word.Shape;
```

Parameters:
- id: number  
  Required. A shape identifier.

Returns: [Word.Shape](/en-us/javascript/api/word/word.shape)

Remarks: [ API set: WordApiDesktop 1.2 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### getByIdOrNullObject(id)

Gets a shape by its identifier. If there isn't a shape with the identifier in this collection, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
getByIdOrNullObject(id: number): Word.Shape;
```

Parameters:
- id: number  
  Required. A shape identifier.

Returns: [Word.Shape](/en-us/javascript/api/word/word.shape)

Remarks: [ API set: WordApiDesktop 1.2 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### getByIds(ids)

Gets the shapes by the identifiers.

```typescript
getByIds(ids: number[]): Word.ShapeCollection;
```

Parameters:
- ids: number[]  
  Required. An array of shape identifiers.

Returns: [Word.ShapeCollection](/en-us/javascript/api/word/word.shapecollection)

Remarks: [ API set: WordApiDesktop 1.2 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### getByNames(names)

Gets the shapes that have the specified names.

```typescript
getByNames(names: string[]): Word.ShapeCollection;
```

Parameters:
- names: string[]  
  Required. An array of shape names.

Returns: [Word.ShapeCollection](/en-us/javascript/api/word/word.shapecollection)

Remarks: [ API set: WordApiDesktop 1.2 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### getByTypes(types)

Gets the shapes that have the specified types.

```typescript
getByTypes(types: Word.ShapeType[]): Word.ShapeCollection;
```

Parameters:
- types: [Word.ShapeType](/en-us/javascript/api/word/word.shapetype)[]  
  Required. An array of shape types.

Returns: [Word.ShapeCollection](/en-us/javascript/api/word/word.shapecollection)

Remarks: [ API set: WordApiDesktop 1.2 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples

```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-shapes-text-boxes.yaml

await Word.run(async (context) => {
  // Sets the properties of the first text box.
  const firstShapeWithTextBox: Word.Shape = context.document.body.shapes
    .getByTypes([Word.ShapeType.textBox])
    .getFirst();
  firstShapeWithTextBox.top = 115;
  firstShapeWithTextBox.left = 0;
  firstShapeWithTextBox.width = 50;
  firstShapeWithTextBox.height = 50;
  await context.sync();

  console.log("The first text box's properties were updated:", firstShapeWithTextBox);
});
```

---

### getFirst()

Gets the first shape in this collection. Throws an `ItemNotFound` error if this collection is empty.

```typescript
getFirst(): Word.Shape;
```

Returns: [Word.Shape](/en-us/javascript/api/word/word.shape)

Remarks: [ API set: WordApiDesktop 1.2 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples

```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-shapes-text-boxes.yaml

await Word.run(async (context) => {
  // Inserts a content control into the first paragraph in the first text box.
  const firstShapeWithTextBox: Word.Shape = context.document.body.shapes
    .getByTypes([Word.ShapeType.textBox])
    .getFirst();
  firstShapeWithTextBox.load("type/body");
  await context.sync();

  const firstParagraphInTextBox: Word.Paragraph = firstShapeWithTextBox.body.paragraphs.getFirst();
  const newControl: Word.ContentControl = firstParagraphInTextBox.insertContentControl();
  newControl.load();
  await context.sync();

  console.log("New content control properties:", newControl);
});
```

---

### getFirstOrNullObject()

Gets the first shape in this collection. If this collection is empty, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
getFirstOrNullObject(): Word.Shape;
```

Returns: [Word.Shape](/en-us/javascript/api/word/word.shape)

Remarks: [ API set: WordApiDesktop 1.2 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### group()

Groups floating shapes in this collection, inline shapes will be skipped. Returns a Shape object that represents the new group of shapes.

```typescript
group(): Word.Shape;
```

Returns: [Word.Shape](/en-us/javascript/api/word/word.shape)

Remarks: [ API set: WordApiDesktop 1.2 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### load(options)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(options?: Word.Interfaces.ShapeCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions): Word.ShapeCollection;
```

Parameters:
- options: [Word.Interfaces.ShapeCollectionLoadOptions](/en-us/javascript/api/word/word.interfaces.shapecollectionloadoptions) & [Word.Interfaces.CollectionLoadOptions](/en-us/javascript/api/word/word.interfaces.collectionloadoptions)  
  Provides options for which properties of the object to load.

Returns: [Word.ShapeCollection](/en-us/javascript/api/word/word.shapecollection)

---

### load(propertyNames)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.ShapeCollection;
```

Parameters:
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns: [Word.ShapeCollection](/en-us/javascript/api/word/word.shapecollection)

---

### load(propertyNamesAndPaths)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Word.ShapeCollection;
```

Parameters:
- propertyNamesAndPaths: [OfficeExtension.LoadOption](/en-us/javascript/api/office/officeextension.loadoption)  
  `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

Returns: [Word.ShapeCollection](/en-us/javascript/api/word/word.shapecollection)

---

### toJSON()

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.ShapeCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.ShapeCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

```typescript
toJSON(): Word.Interfaces.ShapeCollectionData;
```

Returns: [Word.Interfaces.ShapeCollectionData](/en-us/javascript/api/word/word.interfaces.shapecollectiondata)

---

### track()

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.ShapeCollection;
```

Returns: [Word.ShapeCollection](/en-us/javascript/api/word/word.shapecollection)

---

### untrack()

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

```typescript
untrack(): Word.ShapeCollection;
```

Returns: [Word.ShapeCollection](/en-us/javascript/api/word/word.shapecollection)