# Word.Shape class

Package: [word](/en-us/javascript/api/word)

Extends: [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

Represents a shape in the header, footer, or document body. Currently, only the following shapes are supported: text boxes, geometric shapes, groups, pictures, and canvases.

## Remarks

[API set: WordApiDesktop 1.2]

#### Examples

```typescript
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

## Properties

- allowOverlap: Specifies whether a given shape can overlap other shapes.
- altTextDescription: Specifies a string that represents the alternative text associated with the shape.
- body: Represents the body object of the shape. Only applies to text boxes and geometric shapes.
- canvas: Gets the canvas associated with the shape. An object with its isNullObject property set to true will be returned if the shape type isn't "Canvas". For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).
- context: The request context associated with the object. This connects the add-in's process to the Office host application's process.
- fill: Returns the fill formatting of the shape.
- geometricShapeType: The geometric shape type of the shape. It will be null if isn't a geometric shape.
- height: The height, in points, of the shape.
- heightRelative: The percentage of shape height to vertical relative size, see [Word.RelativeSize](/en-us/javascript/api/word/word.relativesize). For an inline or child shape, it can't be set.
- id: Gets an integer that represents the shape identifier.
- isChild: Check whether this shape is a child of a group shape or a canvas shape.
- left: The distance, in points, from the left side of the shape to the horizontal relative position, see [Word.RelativeHorizontalPosition](/en-us/javascript/api/word/word.relativehorizontalposition). For an inline shape, it will return 0 and can't be set. For a child shape in a canvas or group, it's relative to the top left corner.
- leftRelative: The relative left position as a percentage from the left side of the shape to the horizontal relative position, see [Word.RelativeHorizontalPosition](/en-us/javascript/api/word/word.relativehorizontalposition). For an inline or child shape, it will return 0 and can't be set.
- lockAspectRatio: Specifies if the aspect ratio of this shape is locked.
- name: The name of the shape.
- parentCanvas: Gets the top-level parent canvas shape of this child shape. It will be null if it isn't a child shape of a canvas.
- parentGroup: Gets the top-level parent group shape of this child shape. It will be null if it isn't a child shape of a group.
- relativeHorizontalPosition: The relative horizontal position of the shape. For an inline shape, it can't be set. For details, see [Word.RelativeHorizontalPosition](/en-us/javascript/api/word/word.relativehorizontalposition).
- relativeHorizontalSize: The relative horizontal size of the shape. For an inline or child shape, it can't be set. For details, see [Word.RelativeHorizontalPosition](/en-us/javascript/api/word/word.relativehorizontalposition).
- relativeVerticalPosition: The relative vertical position of the shape. For an inline shape, it can't be set. For details, see [Word.RelativeVerticalPosition](/en-us/javascript/api/word/word.relativeverticalposition).
- relativeVerticalSize: The relative vertical size of the shape. For an inline or child shape, it can't be set. For details, see [Word.RelativeHorizontalPosition](/en-us/javascript/api/word/word.relativehorizontalposition).
- rotation: Specifies the rotation, in degrees, of the shape. Not applicable to Canvas shape.
- shapeGroup: Gets the shape group associated with the shape. An object with its isNullObject property set to true will be returned if the shape type isn't "GroupShape". For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).
- textFrame: Gets the text frame object of the shape.
- textWrap: Returns the text wrap formatting of the shape.
- top: The distance, in points, from the top edge of the shape to the vertical relative position (see [Word.RelativeVerticalPosition](/en-us/javascript/api/word/word.relativeverticalposition)). For an inline shape, it will return 0 and can't be set. For a child shape in a canvas or group, it's relative to the top left corner.
- topRelative: The relative top position as a percentage from the top edge of the shape to the vertical relative position, see [Word.RelativeVerticalPosition](/en-us/javascript/api/word/word.relativeverticalposition). For an inline or child shape, it will return 0 and can't be set.
- type: Gets the shape type. Currently, only the following shapes are supported: text boxes, geometric shapes, groups, pictures, and canvases.
- visible: Specifies if the shape is visible. Not applicable to inline shapes.
- width: The width, in points, of the shape.
- widthRelative: The percentage of shape width to horizontal relative size, see [Word.RelativeSize](/en-us/javascript/api/word/word.relativesize). For an inline or child shape, it can't be set.

## Methods

- delete(): Deletes the shape and its content.
- load(options): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNames): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNamesAndPaths): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- moveHorizontally(distance): Moves the shape horizontally by the number of points.
- moveVertically(distance): Moves the shape vertically by the number of points.
- scaleHeight(scaleFactor, scaleType, scaleFrom): Scales the height of the shape by a specified factor. For images, you can indicate whether you want to scale the shape relative to the original or the current size. Shapes other than pictures are always scaled relative to their current height.
- scaleHeight(scaleFactor, scaleType, scaleFrom): Scales the height of the shape by a specified factor. For images, you can indicate whether you want to scale the shape relative to the original or the current size. Shapes other than pictures are always scaled relative to their current height.
- scaleWidth(scaleFactor, scaleType, scaleFrom): Scales the width of the shape by a specified factor. For images, you can indicate whether you want to scale the shape relative to the original or the current size. Shapes other than pictures are always scaled relative to their current height.
- scaleWidth(scaleFactor, scaleType, scaleFrom): Scales the width of the shape by a specified factor. For images, you can indicate whether you want to scale the shape relative to the original or the current size. Shapes other than pictures are always scaled relative to their current height.
- select(selectMultipleShapes): Selects the shape.
- set(properties, options): Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
- set(properties): Sets multiple properties on the object at the same time, based on an existing loaded object.
- toJSON(): Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.Shape object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.ShapeData) that contains shallow copies of any loaded child properties from the original object.
- track(): Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
- untrack(): Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

## Property Details

### allowOverlap

Specifies whether a given shape can overlap other shapes.

```typescript
allowOverlap: boolean;
```

Property Value: `boolean`

Remarks: [API set: WordApiDesktop 1.2]

### altTextDescription

Specifies a string that represents the alternative text associated with the shape.

```typescript
altTextDescription: string;
```

Property Value: `string`

Remarks: [API set: WordApiDesktop 1.2]

### body

Represents the body object of the shape. Only applies to text boxes and geometric shapes.

```typescript
readonly body: Word.Body;
```

Property Value: [Word.Body](/en-us/javascript/api/word/word.body)

Remarks: [API set: WordApiDesktop 1.2]

### canvas

Gets the canvas associated with the shape. An object with its `isNullObject` property set to `true` will be returned if the shape type isn't "Canvas". For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
readonly canvas: Word.Canvas;
```

Property Value: [Word.Canvas](/en-us/javascript/api/word/word.canvas)

Remarks: [API set: WordApiDesktop 1.2]

### context

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value: [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### fill

Returns the fill formatting of the shape.

```typescript
readonly fill: Word.ShapeFill;
```

Property Value: [Word.ShapeFill](/en-us/javascript/api/word/word.shapefill)

Remarks: [API set: WordApiDesktop 1.2]

### geometricShapeType

The geometric shape type of the shape. It will be nul