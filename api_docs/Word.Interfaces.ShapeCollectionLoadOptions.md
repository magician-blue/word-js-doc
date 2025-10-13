# Word.Interfaces.ShapeCollectionLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Contains a collection of [Word.Shape](/en-us/javascript/api/word/word.shape) objects. Currently, only the following shapes are supported: text boxes, geometric shapes, groups, pictures, and canvases.

## Remarks

[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- $all  
  Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).
- allowOverlap  
  For EACH ITEM in the collection: Specifies whether a given shape can overlap other shapes.
- altTextDescription  
  For EACH ITEM in the collection: Specifies a string that represents the alternative text associated with the shape.
- body  
  For EACH ITEM in the collection: Represents the body object of the shape. Only applies to text boxes and geometric shapes.
- canvas  
  For EACH ITEM in the collection: Gets the canvas associated with the shape. An object with its isNullObject property set to true will be returned if the shape type isn't "Canvas". For further information, see [\*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).
- fill  
  For EACH ITEM in the collection: Returns the fill formatting of the shape.
- geometricShapeType  
  For EACH ITEM in the collection: The geometric shape type of the shape. It will be null if isn't a geometric shape.
- height  
  For EACH ITEM in the collection: The height, in points, of the shape.
- heightRelative  
  For EACH ITEM in the collection: The percentage of shape height to vertical relative size, see [Word.RelativeSize](/en-us/javascript/api/word/word.relativesize). For an inline or child shape, it can't be set.
- id  
  For EACH ITEM in the collection: Gets an integer that represents the shape identifier.
- isChild  
  For EACH ITEM in the collection: Check whether this shape is a child of a group shape or a canvas shape.
- left  
  For EACH ITEM in the collection: The distance, in points, from the left side of the shape to the horizontal relative position, see [Word.RelativeHorizontalPosition](/en-us/javascript/api/word/word.relativehorizontalposition). For an inline shape, it will return 0 and can't be set. For a child shape in a canvas or group, it's relative to the top left corner.
- leftRelative  
  For EACH ITEM in the collection: The relative left position as a percentage from the left side of the shape to the horizontal relative position, see [Word.RelativeHorizontalPosition](/en-us/javascript/api/word/word.relativehorizontalposition). For an inline or child shape, it will return 0 and can't be set.
- lockAspectRatio  
  For EACH ITEM in the collection: Specifies if the aspect ratio of this shape is locked.
- name  
  For EACH ITEM in the collection: The name of the shape.
- parentCanvas  
  For EACH ITEM in the collection: Gets the top-level parent canvas shape of this child shape. It will be null if it isn't a child shape of a canvas.
- parentGroup  
  For EACH ITEM in the collection: Gets the top-level parent group shape of this child shape. It will be null if it isn't a child shape of a group.
- relativeHorizontalPosition  
  For EACH ITEM in the collection: The relative horizontal position of the shape. For an inline shape, it can't be set. For details, see [Word.RelativeHorizontalPosition](/en-us/javascript/api/word/word.relativehorizontalposition).
- relativeHorizontalSize  
  For EACH ITEM in the collection: The relative horizontal size of the shape. For an inline or child shape, it can't be set. For details, see [Word.RelativeHorizontalPosition](/en-us/javascript/api/word/word.relativehorizontalposition).
- relativeVerticalPosition  
  For EACH ITEM in the collection: The relative vertical position of the shape. For an inline shape, it can't be set. For details, see [Word.RelativeVerticalPosition](/en-us/javascript/api/word/word.relativeverticalposition).
- relativeVerticalSize  
  For EACH ITEM in the collection: The relative vertical size of the shape. For an inline or child shape, it can't be set. For details, see [Word.RelativeHorizontalPosition](/en-us/javascript/api/word/word.relativehorizontalposition).
- rotation  
  For EACH ITEM in the collection: Specifies the rotation, in degrees, of the shape. Not applicable to Canvas shape.
- shapeGroup  
  For EACH ITEM in the collection: Gets the shape group associated with the shape. An object with its isNullObject property set to true will be returned if the shape type isn't "GroupShape". For further information, see [\*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).
- textFrame  
  For EACH ITEM in the collection: Gets the text frame object of the shape.
- textWrap  
  For EACH ITEM in the collection: Returns the text wrap formatting of the shape.
- top  
  For EACH ITEM in the collection: The distance, in points, from the top edge of the shape to the vertical relative position (see [Word.RelativeVerticalPosition](/en-us/javascript/api/word/word.relativeverticalposition)). For an inline shape, it will return 0 and can't be set. For a child shape in a canvas or group, it's relative to the top left corner.
- topRelative  
  For EACH ITEM in the collection: The relative top position as a percentage from the top edge of the shape to the vertical relative position, see [Word.RelativeVerticalPosition](/en-us/javascript/api/word/word.relativeverticalposition). For an inline or child shape, it will return 0 and can't be set.
- type  
  For EACH ITEM in the collection: Gets the shape type. Currently, only the following shapes are supported: text boxes, geometric shapes, groups, pictures, and canvases.
- visible  
  For EACH ITEM in the collection: Specifies if the shape is visible. Not applicable to inline shapes.
- width  
  For EACH ITEM in the collection: The width, in points, of the shape.
- widthRelative  
  For EACH ITEM in the collection: The percentage of shape width to horizontal relative size, see [Word.RelativeSize](/en-us/javascript/api/word/word.relativesize). For an inline or child shape, it can't be set.

## Property Details

### $all

Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).

```typescript
$all?: boolean;
```

Property Value: boolean

---

### allowOverlap

For EACH ITEM in the collection: Specifies whether a given shape can overlap other shapes.

```typescript
allowOverlap?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### altTextDescription

For EACH ITEM in the collection: Specifies a string that represents the alternative text associated with the shape.

```typescript
altTextDescription?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### body

For EACH ITEM in the collection: Represents the body object of the shape. Only applies to text boxes and geometric shapes.

```typescript
body?: Word.Interfaces.BodyLoadOptions;
```

Property Value: [Word.Interfaces.BodyLoadOptions](/en-us/javascript/api/word/word.interfaces.bodyloadoptions)

Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### canvas

For EACH ITEM in the collection: Gets the canvas associated with the shape. An object with its isNullObject property set to true will be returned if the shape type isn't "Canvas". For further information, see [\*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
canvas?: Word.Interfaces.CanvasLoadOptions;
```

Property Value: [Word.Interfaces.CanvasLoadOptions](/en-us/javascript/api/word/word.interfaces.canvasloadoptions)

Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### fill

For EACH ITEM in the collection: Returns the fill formatting of the shape.

```typescript
fill?: Word.Interfaces.ShapeFillLoadOptions;
```

Property Value: [Word.Interfaces.ShapeFillLoadOptions](/en-us/javascript/api/word/word.interfaces.shapefillloadoptions)

Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### geometricShapeType

For EACH ITEM in the collection: The geometric shape type of the shape. It will be null if isn't a geometric shape.

```typescript
geometricShapeType?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### height

For EACH ITEM in the collection: The height, in points, of the shape.

```typescript
height?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### heightRelative

For EACH ITEM in the collection: The percentage of shape height to vertical relative size, see [Word.RelativeSize](/en-us/javascript/api/word/word.relativesize). For an inline or child shape, it can't be set.

```typescript
heightRelative?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### id

For EACH ITEM in the collection: Gets an integer that represents the shape identifier.

```typescript
id?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### isChild

For EACH ITEM in the collection: Check whether this shape is a child of a group shape or a canvas shape.

```typescript
isChild?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### left

For EACH ITEM in the collection: The distance, in points, from the left side of the shape to the horizontal relative position, see [Word.RelativeHorizontalPosition](/en-us/javascript/api/word/word.relativehorizontalposition). For an inline shape, it will return 0 and can't be set. For a child shape in a canvas or group, it's relative to the top left corner.

```typescript
left?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### leftRelative

For EACH ITEM in the collection: The relative left position as a percentage from the left side of the shape to the horizontal relative position, see [Word.RelativeHorizontalPosition](/en-us/javascript/api/word/word.relativehorizontalposition). For an inline or child shape, it will return 0 and can't be set.

```typescript
leftRelative?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### lockAspectRatio

For EACH ITEM in the collection: Specifies if the aspect ratio of this shape is locked.

```typescript
lockAspectRatio?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### name

For EACH ITEM in the collection: The name of the shape.

```typescript
name?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### parentCanvas

For EACH ITEM in the collection: Gets the top-level parent canvas shape of this child shape. It will be null if it isn't a child shape of a canvas.

```typescript
parentCanvas?: Word.Interfaces.ShapeLoadOptions;
```

Property Value: [Word.Interfaces.ShapeLoadOptions](/en-us/javascript/api/word/word.interfaces.shapeloadoptions)

Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### parentGroup

For EACH ITEM in the collection: Gets the top-level parent group shape of this child shape. It will be null if it isn't a child shape of a group.

```typescript
parentGroup?: Word.Interfaces.ShapeLoadOptions;
```

Property Value: [Word.Interfaces.ShapeLoadOptions](/en-us/javascript/api/word/word.interfaces.shapeloadoptions)

Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### relativeHorizontalPosition

For EACH ITEM in the collection: The relative horizontal position of the shape. For an inline shape, it can't be set. For details, see [Word.RelativeHorizontalPosition](/en-us/javascript/api/word/word.relativehorizontalposition).

```typescript
relativeHorizontalPosition?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### relativeHorizontalSize

For EACH ITEM in the collection: The relative horizontal size of the shape. For an inline or child shape, it can't be set. For details, see [Word.RelativeHorizontalPosition](/en-us/javascript/api/word/word.relativehorizontalposition).

```typescript
relativeHorizontalSize?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### relativeVerticalPosition

For EACH ITEM in the collection: The relative vertical position of the shape. For an inline shape, it can't be set. For details, see [Word.RelativeVerticalPosition](/en-us/javascript/api/word/word.relativeverticalposition).

```typescript
relativeVerticalPosition?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### relativeVerticalSize

For EACH ITEM in the collection: The relative vertical size of the shape. For an inline or child shape, it can't be set. For details, see [Word.RelativeHorizontalPosition](/en-us/javascript/api/word/word.relativehorizontalposition).

```typescript
relativeVerticalSize?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### rotation

For EACH ITEM in the collection: Specifies the rotation, in degrees, of the shape. Not applicable to Canvas shape.

```typescript
rotation?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### shapeGroup

For EACH ITEM in the collection: Gets the shape group associated with the shape. An object with its isNullObject property set to true will be returned if the shape type isn't "GroupShape". For further information, see [\*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
shapeGroup?: Word.Interfaces.ShapeGroupLoadOptions;
```

Property Value: [Word.Interfaces.ShapeGroupLoadOptions](/en-us/javascript/api/word/word.interfaces.shapegrouploadoptions)

Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### textFrame

For EACH ITEM in the collection: Gets the text frame object of the shape.

```typescript
textFrame?: Word.Interfaces.TextFrameLoadOptions;
```

Property Value: [Word.Interfaces.TextFrameLoadOptions](/en-us/javascript/api/word/word.interfaces.textframeloadoptions)

Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### textWrap

For EACH ITEM in the collection: Returns the text wrap formatting of the shape.

```typescript
textWrap?: Word.Interfaces.ShapeTextWrapLoadOptions;
```

Property Value: [Word.Interfaces.ShapeTextWrapLoadOptions](/en-us/javascript/api/word/word.interfaces.shapetextwraploadoptions)

Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### top

For EACH ITEM in the collection: The distance, in points, from the top edge of the shape to the vertical relative position (see [Word.RelativeVerticalPosition](/en-us/javascript/api/word/word.relativeverticalposition)). For an inline shape, it will return 0 and can't be set. For a child shape in a canvas or group, it's relative to the top left corner.

```typescript
top?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### topRelative

For EACH ITEM in the collection: The relative top position as a percentage from the top edge of the shape to the vertical relative position, see [Word.RelativeVerticalPosition](/en-us/javascript/api/word/word.relativeverticalposition). For an inline or child shape, it will return 0 and can't be set.

```typescript
topRelative?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### type

For EACH ITEM in the collection: Gets the shape type. Currently, only the following shapes are supported: text boxes, geometric shapes, groups, pictures, and canvases.

```typescript
type?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### visible

For EACH ITEM in the collection: Specifies if the shape is visible. Not applicable to inline shapes.

```typescript
visible?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### width

For EACH ITEM in the collection: The width, in points, of the shape.

```typescript
width?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### widthRelative

For EACH ITEM in the collection: The percentage of shape width to horizontal relative size, see [Word.RelativeSize](/en-us/javascript/api/word/word.relativesize). For an inline or child shape, it can't be set.

```typescript
widthRelative?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)