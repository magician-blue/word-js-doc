# Word.Interfaces.ShapeUpdateData interface

Package: [word](/en-us/javascript/api/word)

An interface for updating data on the Shape object, for use in shape.set({ ... }).

## Properties

- `allowOverlap` — Specifies whether a given shape can overlap other shapes.
- `altTextDescription` — Specifies a string that represents the alternative text associated with the shape.
- `body` — Represents the body object of the shape. Only applies to text boxes and geometric shapes.
- `canvas` — Gets the canvas associated with the shape. An object with its `isNullObject` property set to `true` will be returned if the shape type isn't "Canvas". For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).
- `fill` — Returns the fill formatting of the shape.
- `geometricShapeType` — The geometric shape type of the shape. It will be null if isn't a geometric shape.
- `height` — The height, in points, of the shape.
- `heightRelative` — The percentage of shape height to vertical relative size, see [Word.RelativeSize](/en-us/javascript/api/word/word.relativesize). For an inline or child shape, it can't be set.
- `left` — The distance, in points, from the left side of the shape to the horizontal relative position, see [Word.RelativeHorizontalPosition](/en-us/javascript/api/word/word.relativehorizontalposition). For an inline shape, it will return 0 and can't be set. For a child shape in a canvas or group, it's relative to the top left corner.
- `leftRelative` — The relative left position as a percentage from the left side of the shape to the horizontal relative position, see [Word.RelativeHorizontalPosition](/en-us/javascript/api/word/word.relativehorizontalposition). For an inline or child shape, it will return 0 and can't be set.
- `lockAspectRatio` — Specifies if the aspect ratio of this shape is locked.
- `name` — The name of the shape.
- `parentCanvas` — Gets the top-level parent canvas shape of this child shape. It will be null if it isn't a child shape of a canvas.
- `parentGroup` — Gets the top-level parent group shape of this child shape. It will be null if it isn't a child shape of a group.
- `relativeHorizontalPosition` — The relative horizontal position of the shape. For an inline shape, it can't be set. For details, see [Word.RelativeHorizontalPosition](/en-us/javascript/api/word/word.relativehorizontalposition).
- `relativeHorizontalSize` — The relative horizontal size of the shape. For an inline or child shape, it can't be set. For details, see [Word.RelativeHorizontalPosition](/en-us/javascript/api/word/word.relativehorizontalposition).
- `relativeVerticalPosition` — The relative vertical position of the shape. For an inline shape, it can't be set. For details, see [Word.RelativeVerticalPosition](/en-us/javascript/api/word/word.relativeverticalposition).
- `relativeVerticalSize` — The relative vertical size of the shape. For an inline or child shape, it can't be set. For details, see [Word.RelativeHorizontalPosition](/en-us/javascript/api/word/word.relativehorizontalposition).
- `rotation` — Specifies the rotation, in degrees, of the shape. Not applicable to Canvas shape.
- `shapeGroup` — Gets the shape group associated with the shape. An object with its `isNullObject` property set to `true` will be returned if the shape type isn't "GroupShape". For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).
- `textFrame` — Gets the text frame object of the shape.
- `textWrap` — Returns the text wrap formatting of the shape.
- `top` — The distance, in points, from the top edge of the shape to the vertical relative position (see [Word.RelativeVerticalPosition](/en-us/javascript/api/word/word.relativeverticalposition)). For an inline shape, it will return 0 and can't be set. For a child shape in a canvas or group, it's relative to the top left corner.
- `topRelative` — The relative top position as a percentage from the top edge of the shape to the vertical relative position, see [Word.RelativeVerticalPosition](/en-us/javascript/api/word/word.relativeverticalposition). For an inline or child shape, it will return 0 and can't be set.
- `visible` — Specifies if the shape is visible. Not applicable to inline shapes.
- `width` — The width, in points, of the shape.
- `widthRelative` — The percentage of shape width to horizontal relative size, see [Word.RelativeSize](/en-us/javascript/api/word/word.relativesize). For an inline or child shape, it can't be set.

## Property Details

### allowOverlap

Specifies whether a given shape can overlap other shapes.

```typescript
allowOverlap?: boolean;
```

#### Property Value
boolean

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### altTextDescription

Specifies a string that represents the alternative text associated with the shape.

```typescript
altTextDescription?: string;
```

#### Property Value
string

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### body

Represents the body object of the shape. Only applies to text boxes and geometric shapes.

```typescript
body?: Word.Interfaces.BodyUpdateData;
```

#### Property Value
[Word.Interfaces.BodyUpdateData](/en-us/javascript/api/word/word.interfaces.bodyupdatedata)

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### canvas

Gets the canvas associated with the shape. An object with its `isNullObject` property set to `true` will be returned if the shape type isn't "Canvas". For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
canvas?: Word.Interfaces.CanvasUpdateData;
```

#### Property Value
[Word.Interfaces.CanvasUpdateData](/en-us/javascript/api/word/word.interfaces.canvasupdatedata)

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### fill

Returns the fill formatting of the shape.

```typescript
fill?: Word.Interfaces.ShapeFillUpdateData;
```

#### Property Value
[Word.Interfaces.ShapeFillUpdateData](/en-us/javascript/api/word/word.interfaces.shapefillupdatedata)

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### geometricShapeType

The geometric shape type of the shape. It will be null if isn't a geometric shape.

```typescript
geometricShapeType?: Word.GeometricShapeType | "LineInverse" | "Triangle" | "RightTriangle" | "Rectangle" | "Diamond" | "Parallelogram" | "Trapezoid" | "NonIsoscelesTrapezoid" | "Pentagon" | "Hexagon" | "Heptagon" | "Octagon" | "Decagon" | "Dodecagon" | "Star4" | "Star5" | "Star6" | "Star7" | "Star8" | "Star10" | "Star12" | "Star16" | "Star24" | "Star32" | "RoundRectangle" | "Round1Rectangle" | "Round2SameRectangle" | "Round2DiagonalRectangle" | "SnipRoundRectangle" | "Snip1Rectangle" | "Snip2SameRectangle" | "Snip2DiagonalRectangle" | "Plaque" | "Ellipse" | "Teardrop" | "HomePlate" | "Chevron" | "PieWedge" | "Pie" | "BlockArc" | "Donut" | "NoSmoking" | "RightArrow" | "LeftArrow" | "UpArrow" | "DownArrow" | "StripedRightArrow" | "NotchedRightArrow" | "BentUpArrow" | "LeftRightArrow" | "UpDownArrow" | "LeftUpArrow" | "LeftRightUpArrow" | "QuadArrow" | "LeftArrowCallout" | "RightArrowCallout" | "UpArrowCallout" | "DownArrowCallout" | "LeftRightArrowCallout" | "UpDownArrowCallout" | "QuadArrowCallout" | "BentArrow" | "UturnArrow" | "CircularArrow" | "LeftCircularArrow" | "LeftRightCircularArrow" | "CurvedRightArrow" | "CurvedLeftArrow" | "CurvedUpArrow" | "CurvedDownArrow" | "SwooshArrow" | "Cube" | "Can" | "LightningBolt" | "Heart" | "Sun" | "Moon" | "SmileyFace" | "IrregularSeal1" | "IrregularSeal2" | "FoldedCorner" | "Bevel" | "Frame" | "HalfFrame" | "Corner" | "DiagonalStripe" | "Chord" | "Arc" | "LeftBracket" | "RightBracket" | "LeftBrace" | "RightBrace" | "BracketPair" | "BracePair" | "Callout1" | "Callout2" | "Callout3" | "AccentCallout1" | "AccentCallout2" | "AccentCallout3" | "BorderCallout1" | "BorderCallout2" | "BorderCallout3" | "AccentBorderCallout1" | "AccentBorderCallout2" | "AccentBorderCallout3" | "WedgeRectCallout" | "WedgeRRectCallout" | "WedgeEllipseCallout" | "CloudCallout" | "Cloud" | "Ribbon" | "Ribbon2" | "EllipseRibbon" | "EllipseRibbon2" | "LeftRightRibbon" | "VerticalScroll" | "HorizontalScroll" | "Wave" | "DoubleWave" | "Plus" | "FlowChartProcess" | "FlowChartDecision" | "FlowChartInputOutput" | "FlowChartPredefinedProcess" | "FlowChartInternalStorage" | "FlowChartDocument" | "FlowChartMultidocument" | "FlowChartTerminator" | "FlowChartPreparation" | "FlowChartManualInput" | "FlowChartManualOperation" | "FlowChartConnector" | "FlowChartPunchedCard" | "FlowChartPunchedTape" | "FlowChartSummingJunction" | "FlowChartOr" | "FlowChartCollate" | "FlowChartSort" | "FlowChartExtract" | "FlowChartMerge" | "FlowChartOfflineStorage" | "FlowChartOnlineStorage" | "FlowChartMagneticTape" | "FlowChartMagneticDisk" | "FlowChartMagneticDrum" | "FlowChartDisplay" | "FlowChartDelay" | "FlowChartAlternateProcess" | "FlowChartOffpageConnector" | "ActionButtonBlank" | "ActionButtonHome" | "ActionButtonHelp" | "ActionButtonInformation" | "ActionButtonForwardNext" | "ActionButtonBackPrevious" | "ActionButtonEnd" | "ActionButtonBeginning" | "ActionButtonReturn" | "ActionButtonDocument" | "ActionButtonSound" | "ActionButtonMovie" | "Gear6" | "Gear9" | "Funnel" | "MathPlus" | "MathMinus" | "MathMultiply" | "MathDivide" | "MathEqual" | "MathNotEqual" | "CornerTabs" | "SquareTabs" | "PlaqueTabs" | "ChartX" | "ChartStar" | "ChartPlus";
```

#### Property Value
[Word.GeometricShapeType](/en-us/javascript/api/word/word.geometricshapetype) | "LineInverse" | "Triangle" | "RightTriangle" | "Rectangle" | "Diamond" | "Parallelogram" | "Trapezoid" | "NonIsoscelesTrapezoid" | "Pentagon" | "Hexagon" | "Heptagon" | "Octagon" | "Decagon" | "Dodecagon" | "Star4" | "Star5" | "Star6" | "Star7" | "Star8" | "Star10" | "Star12" | "Star16" | "Star24" | "Star32" | "RoundRectangle" | "Round1Rectangle" | "Round2SameRectangle" | "Round2DiagonalRectangle" | "SnipRoundRectangle" | "Snip1Rectangle" | "Snip2SameRectangle" | "Snip2DiagonalRectangle" | "Plaque" | "Ellipse" | "Teardrop" | "HomePlate" | "Chevron" | "PieWedge" | "Pie" | "BlockArc" | "Donut" | "NoSmoking" | "RightArrow" | "LeftArrow" | "UpArrow" | "DownArrow" | "StripedRightArrow" | "NotchedRightArrow" | "BentUpArrow" | "LeftRightArrow" | "UpDownArrow" | "LeftUpArrow" | "LeftRightUpArrow" | "QuadArrow" | "LeftArrowCallout" | "RightArrowCallout" | "UpArrowCallout" | "DownArrowCallout" | "LeftRightArrowCallout" | "UpDownArrowCallout" | "QuadArrowCallout" | "BentArrow" | "UturnArrow" | "CircularArrow" | "LeftCircularArrow" | "LeftRightCircularArrow" | "CurvedRightArrow" | "CurvedLeftArrow" | "CurvedUpArrow" | "CurvedDownArrow" | "SwooshArrow" | "Cube" | "Can" | "LightningBolt" | "Heart" | "Sun" | "Moon" | "SmileyFace" | "IrregularSeal1" | "IrregularSeal2" | "FoldedCorner" | "Bevel" | "Frame" | "HalfFrame" | "Corner" | "DiagonalStripe" | "Chord" | "Arc" | "LeftBracket" | "RightBracket" | "LeftBrace" | "RightBrace" | "BracketPair" | "BracePair" | "Callout1" | "Callout2" | "Callout3" | "AccentCallout1" | "AccentCallout2" | "AccentCallout3" | "BorderCallout1" | "BorderCallout2" | "BorderCallout3" | "AccentBorderCallout1" | "AccentBorderCallout2" | "AccentBorderCallout3" | "WedgeRectCallout" | "WedgeRRectCallout" | "WedgeEllipseCallout" | "CloudCallout" | "Cloud" | "Ribbon" | "Ribbon2" | "EllipseRibbon" | "EllipseRibbon2" | "LeftRightRibbon" | "VerticalScroll" | "HorizontalScroll" | "Wave" | "DoubleWave" | "Plus" | "FlowChartProcess" | "FlowChartDecision" | "FlowChartInputOutput" | "FlowChartPredefinedProcess" | "FlowChartInternalStorage" | "FlowChartDocument" | "FlowChartMultidocument" | "FlowChartTerminator" | "FlowChartPreparation" | "FlowChartManualInput" | "FlowChartManualOperation" | "FlowChartConnector" | "FlowChartPunchedCard" | "FlowChartPunchedTape" | "FlowChartSummingJunction" | "FlowChartOr" | "FlowChartCollate" | "FlowChartSort" | "FlowChartExtract" | "FlowChartMerge" | "FlowChartOfflineStorage" | "FlowChartOnlineStorage" | "FlowChartMagneticTape" | "FlowChartMagneticDisk" | "FlowChartMagneticDrum" | "FlowChartDisplay" | "FlowChartDelay" | "FlowChartAlternateProcess" | "FlowChartOffpageConnector" | "ActionButtonBlank" | "ActionButtonHome" | "ActionButtonHelp" | "ActionButtonInformation" | "ActionButtonForwardNext" | "ActionButtonBackPrevious" | "ActionButtonEnd" | "ActionButtonBeginning" | "ActionButtonReturn" | "ActionButtonDocument" | "ActionButtonSound" | "ActionButtonMovie" | "Gear6" | "Gear9" | "Funnel" | "MathPlus" | "MathMinus" | "MathMultiply" | "MathDivide" | "MathEqual" | "MathNotEqual" | "CornerTabs" | "SquareTabs" | "PlaqueTabs" | "ChartX" | "ChartStar" | "ChartPlus"

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### height

The height, in points, of the shape.

```typescript
height?: number;
```

#### Property Value
number

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### heightRelative

The percentage of shape height to vertical relative size, see [Word.RelativeSize](/en-us/javascript/api/word/word.relativesize). For an inline or child shape, it can't be set.

```typescript
heightRelative?: number;
```

#### Property Value
number

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### left

The distance, in points, from the left side of the shape to the horizontal relative position, see [Word.RelativeHorizontalPosition](/en-us/javascript/api/word/word.relativehorizontalposition). For an inline shape, it will return 0 and can't be set. For a child shape in a canvas or group, it's relative to the top left corner.

```typescript
left?: number;
```

#### Property Value
number

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### leftRelative

The relative left position as a percentage from the left side of the shape to the horizontal relative position, see [Word.RelativeHorizontalPosition](/en-us/javascript/api/word/word.relativehorizontalposition). For an inline or child shape, it will return 0 and can't be set.

```typescript
leftRelative?: number;
```

#### Property Value
number

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### lockAspectRatio

Specifies if the aspect ratio of this shape is locked.

```typescript
lockAspectRatio?: boolean;
```

#### Property Value
boolean

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### name

The name of the shape.

```typescript
name?: string;
```

#### Property Value
string

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### parentCanvas

Gets the top-level parent canvas shape of this child shape. It will be null if it isn't a child shape of a canvas.

```typescript
parentCanvas?: Word.Interfaces.ShapeUpdateData;
```

#### Property Value
[Word.Interfaces.ShapeUpdateData](/en-us/javascript/api/word/word.interfaces.shapeupdatedata)

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### parentGroup

Gets the top-level parent group shape of this child shape. It will be null if it isn't a child shape of a group.

```typescript
parentGroup?: Word.Interfaces.ShapeUpdateData;
```

#### Property Value
[Word.Interfaces.ShapeUpdateData](/en-us/javascript/api/word/word.interfaces.shapeupdatedata)

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### relativeHorizontalPosition

The relative horizontal position of the shape. For an inline shape, it can't be set. For details, see [Word.RelativeHorizontalPosition](/en-us/javascript/api/word/word.relativehorizontalposition).

```typescript
relativeHorizontalPosition?: Word.RelativeHorizontalPosition | "Margin" | "Page" | "Column" | "Character" | "LeftMargin" | "RightMargin" | "InsideMargin" | "OutsideMargin";
```

#### Property Value
[Word.RelativeHorizontalPosition](/en-us/javascript/api/word/word.relativehorizontalposition) | "Margin" | "Page" | "Column" | "Character" | "LeftMargin" | "RightMargin" | "InsideMargin" | "OutsideMargin"

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### relativeHorizontalSize

The relative horizontal size of the shape. For an inline or child shape, it can't be set. For details, see [Word.RelativeHorizontalPosition](/en-us/javascript/api/word/word.relativehorizontalposition).

```typescript
relativeHorizontalSize?: Word.RelativeSize | "Margin" | "Page" | "TopMargin" | "BottomMargin" | "InsideMargin" | "OutsideMargin";
```

#### Property Value
[Word.RelativeSize](/en-us/javascript/api/word/word.relativesize) | "Margin" | "Page" | "TopMargin" | "BottomMargin" | "InsideMargin" | "OutsideMargin"

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### relativeVerticalPosition

The relative vertical position of the shape. For an inline shape, it can't be set. For details, see [Word.RelativeVerticalPosition](/en-us/javascript/api/word/word.relativeverticalposition).

```typescript
relativeVerticalPosition?: Word.RelativeVerticalPosition | "Margin" | "Page" | "Paragraph" | "Line" | "TopMargin" | "BottomMargin" | "InsideMargin" | "OutsideMargin";
```

#### Property Value
[Word.RelativeVerticalPosition](/en-us/javascript/api/word/word.relativeverticalposition) | "Margin" | "Page" | "Paragraph" | "Line" | "TopMargin" | "BottomMargin" | "InsideMargin" | "OutsideMargin"

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### relativeVerticalSize

The relative vertical size of the shape. For an inline or child shape, it can't be set. For details, see [Word.RelativeHorizontalPosition](/en-us/javascript/api/word/word.relativehorizontalposition).

```typescript
relativeVerticalSize?: Word.RelativeSize | "Margin" | "Page" | "TopMargin" | "BottomMargin" | "InsideMargin" | "OutsideMargin";
```

#### Property Value
[Word.RelativeSize](/en-us/javascript/api/word/word.relativesize) | "Margin" | "Page" | "TopMargin" | "BottomMargin" | "InsideMargin" | "OutsideMargin"

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### rotation

Specifies the rotation, in degrees, of the shape. Not applicable to Canvas shape.

```typescript
rotation?: number;
```

#### Property Value
number

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### shapeGroup

Gets the shape group associated with the shape. An object with its `isNullObject` property set to `true` will be returned if the shape type isn't "GroupShape". For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
shapeGroup?: Word.Interfaces.ShapeGroupUpdateData;
```

#### Property Value
[Word.Interfaces.ShapeGroupUpdateData](/en-us/javascript/api/word/word.interfaces.shapegroupupdatedata)

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### textFrame

Gets the text frame object of the shape.

```typescript
textFrame?: Word.Interfaces.TextFrameUpdateData;
```

#### Property Value
[Word.Interfaces.TextFrameUpdateData](/en-us/javascript/api/word/word.interfaces.textframeupdatedata)

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### textWrap

Returns the text wrap formatting of the shape.

```typescript
textWrap?: Word.Interfaces.ShapeTextWrapUpdateData;
```

#### Property Value
[Word.Interfaces.ShapeTextWrapUpdateData](/en-us/javascript/api/word/word.interfaces.shapetextwrapupdatedata)

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### top

The distance, in points, from the top edge of the shape to the vertical relative position (see [Word.RelativeVerticalPosition](/en-us/javascript/api/word/word.relativeverticalposition)). For an inline shape, it will return 0 and can't be set. For a child shape in a canvas or group, it's relative to the top left corner.

```typescript
top?: number;
```

#### Property Value
number

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### topRelative

The relative top position as a percentage from the top edge of the shape to the vertical relative position, see [Word.RelativeVerticalPosition](/en-us/javascript/api/word/word.relativeverticalposition). For an inline or child shape, it will return 0 and can't be set.

```typescript
topRelative?: number;
```

#### Property Value
number

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### visible

Specifies if the shape is visible. Not applicable to inline shapes.

```typescript
visible?: boolean;
```

#### Property Value
boolean

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### width

The width, in points, of the shape.

```typescript
width?: number;
```

#### Property Value
number

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### widthRelative

The percentage of shape width to horizontal relative size, see [Word.RelativeSize](/en-us/javascript/api/word/word.relativesize). For an inline or child shape, it can't be set.

```typescript
widthRelative?: number;
```

#### Property Value
number

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)