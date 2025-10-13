# Word.Interfaces.ShapeData interface

Package: [word](/en-us/javascript/api/word)

An interface describing the data returned by calling `shape.toJSON()`.

## Properties

- [allowOverlap](#allowoverlap) — Specifies whether a given shape can overlap other shapes.
- [altTextDescription](#alttextdescription) — Specifies a string that represents the alternative text associated with the shape.
- [body](#body) — Represents the body object of the shape. Only applies to text boxes and geometric shapes.
- [canvas](#canvas) — Gets the canvas associated with the shape. An object with its `isNullObject` property set to `true` will be returned if the shape type isn't "Canvas". For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).
- [fill](#fill) — Returns the fill formatting of the shape.
- [geometricShapeType](#geometricshapetype) — The geometric shape type of the shape. It will be null if isn't a geometric shape.
- [height](#height) — The height, in points, of the shape.
- [heightRelative](#heightrelative) — The percentage of shape height to vertical relative size, see [Word.RelativeSize](/en-us/javascript/api/word/word.relativesize). For an inline or child shape, it can't be set.
- [id](#id) — Gets an integer that represents the shape identifier.
- [isChild](#ischild) — Check whether this shape is a child of a group shape or a canvas shape.
- [left](#left) — The distance, in points, from the left side of the shape to the horizontal relative position, see [Word.RelativeHorizontalPosition](/en-us/javascript/api/word/word.relativehorizontalposition). For an inline shape, it will return 0 and can't be set. For a child shape in a canvas or group, it's relative to the top left corner.
- [leftRelative](#leftrelative) — The relative left position as a percentage from the left side of the shape to the horizontal relative position, see [Word.RelativeHorizontalPosition](/en-us/javascript/api/word/word.relativehorizontalposition). For an inline or child shape, it will return 0 and can't be set.
- [lockAspectRatio](#lockaspectratio) — Specifies if the aspect ratio of this shape is locked.
- [name](#name) — The name of the shape.
- [parentCanvas](#parentcanvas) — Gets the top-level parent canvas shape of this child shape. It will be null if it isn't a child shape of a canvas.
- [parentGroup](#parentgroup) — Gets the top-level parent group shape of this child shape. It will be null if it isn't a child shape of a group.
- [relativeHorizontalPosition](#relativehorizontalposition) — The relative horizontal position of the shape. For an inline shape, it can't be set. For details, see [Word.RelativeHorizontalPosition](/en-us/javascript/api/word/word.relativehorizontalposition).
- [relativeHorizontalSize](#relativehorizontalsize) — The relative horizontal size of the shape. For an inline or child shape, it can't be set. For details, see [Word.RelativeHorizontalPosition](/en-us/javascript/api/word/word.relativehorizontalposition).
- [relativeVerticalPosition](#relativeverticalposition) — The relative vertical position of the shape. For an inline shape, it can't be set. For details, see [Word.RelativeVerticalPosition](/en-us/javascript/api/word/word.relativeverticalposition).
- [relativeVerticalSize](#relativeverticalsize) — The relative vertical size of the shape. For an inline or child shape, it can't be set. For details, see [Word.RelativeHorizontalPosition](/en-us/javascript/api/word/word.relativehorizontalposition).
- [rotation](#rotation) — Specifies the rotation, in degrees, of the shape. Not applicable to Canvas shape.
- [shapeGroup](#shapegroup) — Gets the shape group associated with the shape. An object with its `isNullObject` property set to `true` will be returned if the shape type isn't "GroupShape". For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).
- [textFrame](#textframe) — Gets the text frame object of the shape.
- [textWrap](#textwrap) — Returns the text wrap formatting of the shape.
- [top](#top) — The distance, in points, from the top edge of the shape to the vertical relative position (see [Word.RelativeVerticalPosition](/en-us/javascript/api/word/word.relativeverticalposition)). For an inline shape, it will return 0 and can't be set. For a child shape in a canvas or group, it's relative to the top left corner.
- [topRelative](#toprelative) — The relative top position as a percentage from the top edge of the shape to the vertical relative position, see [Word.RelativeVerticalPosition](/en-us/javascript/api/word/word.relativeverticalposition). For an inline or child shape, it will return 0 and can't be set.
- [type](#type) — Gets the shape type. Currently, only the following shapes are supported: text boxes, geometric shapes, groups, pictures, and canvases.
- [visible](#visible) — Specifies if the shape is visible. Not applicable to inline shapes.
- [width](#width) — The width, in points, of the shape.
- [widthRelative](#widthrelative) — The percentage of shape width to horizontal relative size, see [Word.RelativeSize](/en-us/javascript/api/word/word.relativesize). For an inline or child shape, it can't be set.

## Property Details

### allowOverlap

Specifies whether a given shape can overlap other shapes.

```typescript
allowOverlap?: boolean;
```

#### Property value
boolean

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### altTextDescription

Specifies a string that represents the alternative text associated with the shape.

```typescript
altTextDescription?: string;
```

#### Property value
string

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### body

Represents the body object of the shape. Only applies to text boxes and geometric shapes.

```typescript
body?: Word.Interfaces.BodyData;
```

#### Property value
[Word.Interfaces.BodyData](/en-us/javascript/api/word/word.interfaces.bodydata)

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### canvas

Gets the canvas associated with the shape. An object with its `isNullObject` property set to `true` will be returned if the shape type isn't "Canvas". For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
canvas?: Word.Interfaces.CanvasData;
```

#### Property value
[Word.Interfaces.CanvasData](/en-us/javascript/api/word/word.interfaces.canvasdata)

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### fill

Returns the fill formatting of the shape.

```typescript
fill?: Word.Interfaces.ShapeFillData;
```

#### Property value
[Word.Interfaces.ShapeFillData](/en-us/javascript/api/word/word.interfaces.shapefilldata)

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### geometricShapeType

The geometric shape type of the shape. It will be null if isn't a geometric shape.

```typescript
geometricShapeType?: Word.GeometricShapeType | "LineInverse" | "Triangle" | "RightTriangle" | "Rectangle" | "Diamond" | "Parallelogram" | "Trapezoid" | "NonIsoscelesTrapezoid" | "Pentagon" | "Hexagon" | "Heptagon" | "Octagon" | "Decagon" | "Dodecagon" | "Star4" | "Star5" | "Star6" | "Star7" | "Star8" | "Star10" | "Star12" | "Star16" | "Star24" | "Star32" | "RoundRectangle" | "Round1Rectangle" | "Round2SameRectangle" | "Round2DiagonalRectangle" | "SnipRoundRectangle" | "Snip1Rectangle" | "Snip2SameRectangle" | "Snip2DiagonalRectangle" | "Plaque" | "Ellipse" | "Teardrop" | "HomePlate" | "Chevron" | "PieWedge" | "Pie" | "BlockArc" | "Donut" | "NoSmoking" | "RightArrow" | "LeftArrow" | "UpArrow" | "DownArrow" | "StripedRightArrow" | "NotchedRightArrow" | "BentUpArrow" | "LeftRightArrow" | "UpDownArrow" | "LeftUpArrow" | "LeftRightUpArrow" | "QuadArrow" | "LeftArrowCallout" | "RightArrowCallout" | "UpArrowCallout" | "DownArrowCallout" | "LeftRightArrowCallout" | "UpDownArrowCallout" | "QuadArrowCallout" | "BentArrow" | "UturnArrow" | "CircularArrow" | "LeftCircularArrow" | "LeftRightCircularArrow" | "CurvedRightArrow" | "CurvedLeftArrow" | "CurvedUpArrow" | "CurvedDownArrow" | "SwooshArrow" | "Cube" | "Can" | "LightningBolt" | "Heart" | "Sun" | "Moon" | "SmileyFace" | "IrregularSeal1" | "IrregularSeal2" | "FoldedCorner" | "Bevel" | "Frame" | "HalfFrame" | "Corner" | "DiagonalStripe" | "Chord" | "Arc" | "LeftBracket" | "RightBracket" | "LeftBrace" | "RightBrace" | "BracketPair" | "BracePair" | "Callout1" | "Callout2" | "Callout3" | "AccentCallout1" | "AccentCallout2" | "AccentCallout3" | "BorderCallout1" | "BorderCallout2" | "BorderCallout3" | "AccentBorderCallout1" | "AccentBorderCallout2" | "AccentBorderCallout3" | "WedgeRectCallout" | "WedgeRRectCallout" | "WedgeEllipseCallout" | "CloudCallout" | "Cloud" | "Ribbon" | "Ribbon2" | "EllipseRibbon" | "EllipseRibbon2" | "LeftRightRibbon" | "VerticalScroll" | "HorizontalScroll" | "Wave" | "DoubleWave" | "Plus" | "FlowChartProcess" | "FlowChartDecision" | "FlowChartInputOutput" | "FlowChartPredefinedProcess" | "FlowChartInternalStorage" | "FlowChartDocument" | "FlowChartMultidocument" | "FlowChartTerminator" | "FlowChartPreparation" | "FlowChartManualInput" | "FlowChartManualOperation" | "FlowChartConnector" | "FlowChartPunchedCard" | "FlowChartPunchedTape" | "FlowChartSummingJunction" | "FlowChartOr" | "FlowChartCollate" | "FlowChartSort" | "FlowChartExtract" | "FlowChartMerge" | "FlowChartOfflineStorage" | "FlowChartOnlineStorage" | "FlowChartMagneticTape" | "FlowChartMagneticDisk" | "FlowChartMagneticDrum" | "FlowChartDisplay" | "FlowChartDelay" | "FlowChartAlternateProcess" | "FlowChartOffpageConnector" | "ActionButtonBlank" | "ActionButtonHome" | "ActionButtonHelp" | "ActionButtonInformation" | "ActionButtonForwardNext" | "ActionButtonBackPrevious" | "ActionButtonEnd" | "ActionButtonBeginning" | "ActionButtonReturn" | "ActionButtonDocument" | "ActionButtonSound" | "ActionButtonMovie" | "Gear6" | "Gear9" | "Funnel" | "MathPlus" | "MathMinus" | "MathMultiply" | "MathDivide" | "MathEqual" | "MathNotEqual" | "CornerTabs" | "SquareTabs" | "PlaqueTabs" | "ChartX" | "ChartStar" | "ChartPlus";
```

#### Property value
[Word.GeometricShapeType](/en-us/javascript/api/word/word.geometricshapetype) | "LineInverse" | "Triangle" | "RightTriangle" | "Rectangle" | "Diamond" | "Parallelogram" | "Trapezoid" | "NonIsoscelesTrapezoid" | "Pentagon" | "Hexagon" | "Heptagon" | "Octagon" | "Decagon" | "Dodecagon" | "Star4" | "Star5" | "Star6" | "Star7" | "Star8" | "Star10" | "Star12" | "Star16" | "Star24" | "Star32" | "RoundRectangle" | "Round1Rectangle" | "Round2SameRectangle" | "Round2DiagonalRectangle" | "SnipRoundRectangle" | "Snip1Rectangle" | "Snip2SameRectangle" | "Snip2DiagonalRectangle" | "Plaque" | "Ellipse" | "Teardrop" | "HomePlate" | "Chevron" | "PieWedge" | "Pie" | "BlockArc" | "Donut" | "NoSmoking" | "RightArrow" | "LeftArrow" | "UpArrow" | "DownArrow" | "StripedRightArrow" | "NotchedRightArrow" | "BentUpArrow" | "LeftRightArrow" | "UpDownArrow" | "LeftUpArrow" | "LeftRightUpArrow" | "QuadArrow" | "LeftArrowCallout" | "RightArrowCallout" | "UpArrowCallout" | "DownArrowCallout" | "LeftRightArrowCallout" | "UpDownArrowCallout" | "QuadArrowCallout" | "BentArrow" | "UturnArrow" | "CircularArrow" | "LeftCircularArrow" | "LeftRightCircularArrow" | "CurvedRightArrow" | "CurvedLeftArrow" | "CurvedUpArrow" | "CurvedDownArrow" | "SwooshArrow" | "Cube" | "Can" | "LightningBolt" | "Heart" | "Sun" | "Moon" | "SmileyFace" | "IrregularSeal1" | "IrregularSeal2" | "FoldedCorner" | "Bevel" | "Frame" | "HalfFrame" | "Corner" | "DiagonalStripe" | "Chord" | "Arc" | "LeftBracket" | "RightBracket" | "LeftBrace" | "RightBrace" | "BracketPair" | "BracePair" | "Callout1" | "Callout2" | "Callout3" | "AccentCallout1" | "AccentCallout2" | "AccentCallout3" | "BorderCallout1" | "BorderCallout2" | "BorderCallout3" | "AccentBorderCallout1" | "AccentBorderCallout2" | "AccentBorderCallout3" | "WedgeRectCallout" | "WedgeRRectCallout" | "WedgeEllipseCallout" | "CloudCallout" | "Cloud" | "Ribbon" | "Ribbon2" | "EllipseRibbon" | "EllipseRibbon2" | "LeftRightRibbon" | "VerticalScroll" | "HorizontalScroll" | "Wave" | "DoubleWave" | "Plus" | "FlowChartProcess" | "FlowChartDecision" | "FlowChartInputOutput" | "FlowChartPredefinedProcess" | "FlowChartInternalStorage" | "FlowChartDocument" | "FlowChartMultidocument" | "FlowChartTerminator" | "FlowChartPreparation" | "FlowChartManualInput" | "FlowChartManualOperation" | "FlowChartConnector" | "FlowChartPunchedCard" | "FlowChartPunchedTape" | "FlowChartSummingJunction" | "FlowChartOr" | "FlowChartCollate" | "FlowChartSort" | "FlowChartExtract" | "FlowChartMerge" | "FlowChartOfflineStorage" | "FlowChartOnlineStorage" | "FlowChartMagneticTape" | "FlowChartMagneticDisk" | "FlowChartMagneticDrum" | "FlowChartDisplay" | "FlowChartDelay" | "FlowChartAlternateProcess" | "FlowChartOffpageConnector" | "ActionButtonBlank" | "ActionButtonHome" | "ActionButtonHelp" | "ActionButtonInformation" | "ActionButtonForwardNext" | "ActionButtonBackPrevious" | "ActionButtonEnd" | "ActionButtonBeginning" | "ActionButtonReturn" | "ActionButtonDocument" | "ActionButtonSound" | "ActionButtonMovie" | "Gear6" | "Gear9" | "Funnel" | "MathPlus" | "MathMinus" | "MathMultiply" | "MathDivide" | "MathEqual" | "MathNotEqual" | "CornerTabs" | "SquareTabs" | "PlaqueTabs" | "ChartX" | "ChartStar" | "ChartPlus"

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### height

The height, in points, of the shape.

```typescript
height?: number;
```

#### Property value
number

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### heightRelative

The percentage of shape height to vertical relative size, see [Word.RelativeSize](/en-us/javascript/api/word/word.relativesize). For an inline or child shape, it can't be set.

```typescript
heightRelative?: number;
```

#### Property value
number

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### id

Gets an integer that represents the shape identifier.

```typescript
id?: number;
```

#### Property value
number

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### isChild

Check whether this shape is a child of a group shape or a canvas shape.

```typescript
isChild?: boolean;
```

#### Property value
boolean

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### left

The distance, in points, from the left side of the shape to the horizontal relative position, see [Word.RelativeHorizontalPosition](/en-us/javascript/api/word/word.relativehorizontalposition). For an inline shape, it will return 0 and can't be set. For a child shape in a canvas or group, it's relative to the top left corner.

```typescript
left?: number;
```

#### Property value
number

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### leftRelative

The relative left position as a percentage from the left side of the shape to the horizontal relative position, see [Word.RelativeHorizontalPosition](/en-us/javascript/api/word/word.relativehorizontalposition). For an inline or child shape, it will return 0 and can't be set.

```typescript
leftRelative?: number;
```

#### Property value
number

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### lockAspectRatio

Specifies if the aspect ratio of this shape is locked.

```typescript
lockAspectRatio?: boolean;
```

#### Property value
boolean

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### name

The name of the shape.

```typescript
name?: string;
```

#### Property value
string

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### parentCanvas

Gets the top-level parent canvas shape of this child shape. It will be null if it isn't a child shape of a canvas.

```typescript
parentCanvas?: Word.Interfaces.ShapeData;
```

#### Property value
[Word.Interfaces.ShapeData](/en-us/javascript/api/word/word.interfaces.shapedata)

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### parentGroup

Gets the top-level parent group shape of this child shape. It will be null if it isn't a child shape of a group.

```typescript
parentGroup?: Word.Interfaces.ShapeData;
```

#### Property value
[Word.Interfaces.ShapeData](/en-us/javascript/api/word/word.interfaces.shapedata)

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### relativeHorizontalPosition

The relative horizontal position of the shape. For an inline shape, it can't be set. For details, see [Word.RelativeHorizontalPosition](/en-us/javascript/api/word/word.relativehorizontalposition).

```typescript
relativeHorizontalPosition?: Word.RelativeHorizontalPosition | "Margin" | "Page" | "Column" | "Character" | "LeftMargin" | "RightMargin" | "InsideMargin" | "OutsideMargin";
```

#### Property value
[Word.RelativeHorizontalPosition](/en-us/javascript/api/word/word.relativehorizontalposition) | "Margin" | "Page" | "Column" | "Character" | "LeftMargin" | "RightMargin" | "InsideMargin" | "OutsideMargin"

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### relativeHorizontalSize

The relative horizontal size of the shape. For an inline or child shape, it can't be set. For details, see [Word.RelativeHorizontalPosition](/en-us/javascript/api/word/word.relativehorizontalposition).

```typescript
relativeHorizontalSize?: Word.RelativeSize | "Margin" | "Page" | "TopMargin" | "BottomMargin" | "InsideMargin" | "OutsideMargin";
```

#### Property value
[Word.RelativeSize](/en-us/javascript/api/word/word.relativesize) | "Margin" | "Page" | "TopMargin" | "BottomMargin" | "InsideMargin" | "OutsideMargin"

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### relativeVerticalPosition

The relative vertical position of the shape. For an inline shape, it can't be set. For details, see [Word.RelativeVerticalPosition](/en-us/javascript/api/word/word.relativeverticalposition).

```typescript
relativeVerticalPosition?: Word.RelativeVerticalPosition | "Margin" | "Page" | "Paragraph" | "Line" | "TopMargin" | "BottomMargin" | "InsideMargin" | "OutsideMargin";
```

#### Property value
[Word.RelativeVerticalPosition](/en-us/javascript/api/word/word.relativeverticalposition) | "Margin" | "Page" | "Paragraph" | "Line" | "TopMargin" | "BottomMargin" | "InsideMargin" | "OutsideMargin"

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### relativeVerticalSize

The relative vertical size of the shape. For an inline or child shape, it can't be set. For details, see [Word.RelativeHorizontalPosition](/en-us/javascript/api/word/word.relativehorizontalposition).

```typescript
relativeVerticalSize?: Word.RelativeSize | "Margin" | "Page" | "TopMargin" | "BottomMargin" | "InsideMargin" | "OutsideMargin";
```

#### Property value
[Word.RelativeSize](/en-us/javascript/api/word/word.relativesize) | "Margin" | "Page" | "TopMargin" | "BottomMargin" | "InsideMargin" | "OutsideMargin"

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### rotation

Specifies the rotation, in degrees, of the shape. Not applicable to Canvas shape.

```typescript
rotation?: number;
```

#### Property value
number

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### shapeGroup

Gets the shape group associated with the shape. An object with its `isNullObject` property set to `true` will be returned if the shape type isn't "GroupShape". For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
shapeGroup?: Word.Interfaces.ShapeGroupData;
```

#### Property value
[Word.Interfaces.ShapeGroupData](/en-us/javascript/api/word/word.interfaces.shapegroupdata)

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### textFrame

Gets the text frame object of the shape.

```typescript
textFrame?: Word.Interfaces.TextFrameData;
```

#### Property value
[Word.Interfaces.TextFrameData](/en-us/javascript/api/word/word.interfaces.textframedata)

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### textWrap

Returns the text wrap formatting of the shape.

```typescript
textWrap?: Word.Interfaces.ShapeTextWrapData;
```

#### Property value
[Word.Interfaces.ShapeTextWrapData](/en-us/javascript/api/word/word.interfaces.shapetextwrapdata)

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### top

The distance, in points, from the top edge of the shape to the vertical relative position (see [Word.RelativeVerticalPosition](/en-us/javascript/api/word/word.relativeverticalposition)). For an inline shape, it will return 0 and can't be set. For a child shape in a canvas or group, it's relative to the top left corner.

```typescript
top?: number;
```

#### Property value
number

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### topRelative

The relative top position as a percentage from the top edge of the shape to the vertical relative position, see [Word.RelativeVerticalPosition](/en-us/javascript/api/word/word.relativeverticalposition). For an inline or child shape, it will return 0 and can't be set.

```typescript
topRelative?: number;
```

#### Property value
number

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### type

Gets the shape type. Currently, only the following shapes are supported: text boxes, geometric shapes, groups, pictures, and canvases.

```typescript
type?: Word.ShapeType | "Unsupported" | "TextBox" | "GeometricShape" | "Group" | "Picture" | "Canvas";
```

#### Property value
[Word.ShapeType](/en-us/javascript/api/word/word.shapetype) | "Unsupported" | "TextBox" | "GeometricShape" | "Group" | "Picture" | "Canvas"

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### visible

Specifies if the shape is visible. Not applicable to inline shapes.

```typescript
visible?: boolean;
```

#### Property value
boolean

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### width

The width, in points, of the shape.

```typescript
width?: number;
```

#### Property value
number

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### widthRelative

The percentage of shape width to horizontal relative size, see [Word.RelativeSize](/en-us/javascript/api/word/word.relativesize). For an inline or child shape, it can't be set.

```typescript
widthRelative?: number;
```

#### Property value
number

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)