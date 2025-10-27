# LineFormat

**Package:** `word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents line and arrowhead formatting. For a line, the LineFormat object contains formatting information for the line itself; for a shape with a border, this object contains formatting information for the shape's border.

## Properties

### backgroundColor

**Type:** `Word.ColorFormat`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets a ColorFormat object that represents the background color for a patterned line.

#### Examples

**Example**: Set the background color of a patterned line to light blue

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.getShapes();
    const shape = shapes.getFirst();
    
    // Access the line format
    const lineFormat = shape.line;
    
    // Set the background color for the patterned line to light blue
    lineFormat.backgroundColor.set("#ADD8E6");
    
    await context.sync();
});
```

---

### beginArrowheadLength

**Type:** `Word.ArrowheadLength | "Mixed" | "Short" | "Medium" | "Long"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the length of the arrowhead at the beginning of the line.

#### Examples

**Example**: Set the beginning arrowhead of a line shape to have a long length

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document (assuming it's a line)
    const shapes = context.document.body.shapes;
    const line = shapes.getFirst();
    
    // Set the beginning arrowhead length to long
    line.lineFormat.beginArrowheadLength = Word.ArrowheadLength.long;
    
    await context.sync();
});
```

---

### beginArrowheadStyle

**Type:** `Word.ArrowheadStyle | "Mixed" | "None" | "Triangle" | "Open" | "Stealth" | "Diamond" | "Oval"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the style of the arrowhead at the beginning of the line.

#### Examples

**Example**: Set the beginning arrowhead style of a line shape to a triangle arrow

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document (assuming it's a line)
    const shape = context.document.body.shapes.getFirst();
    
    // Set the beginning arrowhead style to triangle
    shape.lineFormat.beginArrowheadStyle = Word.ArrowheadStyle.triangle;
    
    await context.sync();
    
    console.log("Beginning arrowhead style set to triangle");
});
```

---

### beginArrowheadWidth

**Type:** `Word.ArrowheadWidth | "Mixed" | "Narrow" | "Medium" | "Wide"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the width of the arrowhead at the beginning of the line.

#### Examples

**Example**: Set the beginning arrowhead width of a line shape to "Wide"

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.shapes;
    const line = shapes.getItem(0); // Get the first shape (assumed to be a line)
    
    line.lineFormat.beginArrowheadWidth = Word.ArrowheadWidth.wide;
    
    await context.sync();
});
```

---

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a LineFormat object to verify the connection between the add-in and Word host application

```typescript
await Word.run(async (context) => {
    // Get a shape from the document
    const shapes = context.document.body.getShapes();
    const shape = shapes.getFirst();
    const lineFormat = shape.line;
    
    // Load the line format
    lineFormat.load("weight");
    await context.sync();
    
    // Access the request context from the LineFormat object
    const lineContext = lineFormat.context;
    
    // Verify the context is connected (both should reference the same context)
    console.log("Contexts match:", lineContext === context);
    console.log("Line weight:", lineFormat.weight);
});
```

---

### dashStyle

**Type:** `Word.LineDashStyle | "Mixed" | "Solid" | "SquareDot" | "RoundDot" | "Dash" | "DashDot" | "DashDotDot" | "LongDash" | "LongDashDot" | "LongDashDotDot" | "SysDash" | "SysDot" | "SysDashDot"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the dash style for the line.

#### Examples

**Example**: Set a shape's border to use a dash-dot line style instead of a solid line

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    const shape = shapes.getFirst();
    
    // Set the border line style to dash-dot
    shape.lineFormat.dashStyle = Word.LineDashStyle.dashDot;
    
    await context.sync();
});
```

---

### endArrowheadLength

**Type:** `Word.ArrowheadLength | "Mixed" | "Short" | "Medium" | "Long"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the length of the arrowhead at the end of the line.

#### Examples

**Example**: Set the arrowhead at the end of a line shape to be long in length

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document (assuming it's a line)
    const shape = context.document.body.shapes.getFirst();
    
    // Set the end arrowhead length to long
    shape.lineFormat.endArrowheadLength = Word.ArrowheadLength.long;
    
    await context.sync();
});
```

---

### endArrowheadStyle

**Type:** `Word.ArrowheadStyle | "Mixed" | "None" | "Triangle" | "Open" | "Stealth" | "Diamond" | "Oval"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the style of the arrowhead at the end of the line.

#### Examples

**Example**: Set the arrowhead at the end of a line shape to a triangle style

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document (assuming it's a line)
    const shape = context.document.body.shapes.getFirst();
    
    // Set the end arrowhead style to triangle
    shape.lineFormat.endArrowheadStyle = Word.ArrowheadStyle.triangle;
    
    await context.sync();
    
    console.log("End arrowhead style set to triangle");
});
```

---

### endArrowheadWidth

**Type:** `Word.ArrowheadWidth | "Mixed" | "Narrow" | "Medium" | "Wide"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the width of the arrowhead at the end of the line.

#### Examples

**Example**: Set the arrowhead width at the end of a line shape to wide

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document (assuming it's a line)
    const shapes = context.document.body.shapes;
    const line = shapes.getFirst();
    
    // Set the end arrowhead width to wide
    line.lineFormat.endArrowheadWidth = Word.ArrowheadWidth.wide;
    
    await context.sync();
});
```

---

### foregroundColor

**Type:** `Word.ColorFormat`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets a ColorFormat object that represents the foreground color for the line.

#### Examples

**Example**: Set a shape's border color to red by accessing its line format's foreground color

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    const shape = shapes.getFirst();
    
    // Access the line format and set the foreground color to red
    const lineFormat = shape.lineFormat;
    lineFormat.foregroundColor.set("#FF0000");
    
    await context.sync();
});
```

---

### insetPen

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies if to draw lines inside a shape.

#### Examples

**Example**: Configure a rectangle shape to draw its border lines inside the shape boundaries rather than centered on the edge

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    // Access the first shape and set insetPen to true
    const shape = shapes.items[0];
    shape.lineFormat.insetPen = true;
    
    await context.sync();
});
```

---

### isVisible

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies if the object, or the formatting applied to it, is visible.

#### Examples

**Example**: Hide the border of a rectangle shape by setting its line format visibility to false

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    const shape = shapes.getFirst();
    
    // Get the line format and set visibility to false
    const lineFormat = shape.lineFormat;
    lineFormat.isVisible = false;
    
    await context.sync();
});
```

---

### pattern

**Type:** `Word.PatternType | "Mixed" | "Percent5" | "Percent10" | "Percent20" | "Percent25" | "Percent30" | "Percent40" | "Percent50" | "Percent60" | "Percent70" | "Percent75" | "Percent80" | "Percent90" | "DarkHorizontal" | "DarkVertical" | "DarkDownwardDiagonal" | "DarkUpwardDiagonal" | "SmallCheckerBoard" | "Trellis" | "LightHorizontal" | "LightVertical" | "LightDownwardDiagonal" | "LightUpwardDiagonal" | "SmallGrid" | "DottedDiamond" | "WideDownwardDiagonal" | "WideUpwardDiagonal" | "DashedUpwardDiagonal" | "DashedDownwardDiagonal" | "NarrowVertical" | "NarrowHorizontal" | "DashedVertical" | "DashedHorizontal" | "LargeConfetti" | "LargeGrid" | "HorizontalBrick" | "LargeCheckerBoard" | "SmallConfetti" | "ZigZag" | "SolidDiamond" | "DiagonalBrick" | "OutlinedDiamond" | "Plaid" | "Sphere" | "Weave" | "DottedGrid" | "Divot" | "Shingle" | "Wave" | "Horizontal" | "Vertical" | "Cross" | "DownwardDiagonal" | "UpwardDiagonal" | "DiagonalCross"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the pattern applied to the line.

#### Examples

**Example**: Apply a diagonal cross pattern to a shape's border line

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        
        // Set the line pattern to diagonal cross
        shape.lineFormat.pattern = "DiagonalCross";
        
        await context.sync();
        console.log("Shape border pattern set to DiagonalCross");
    }
});
```

---

### style

**Type:** `Word.LineFormatStyle | "Mixed" | "Single" | "ThinThin" | "ThinThick" | "ThickThin" | "ThickBetweenThin"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the line format style.

#### Examples

**Example**: Set a shape's border line style to a thick-thin double line format

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    const shape = shapes.getFirst();
    
    // Set the line format style to ThickThin
    shape.lineFormat.style = Word.LineFormatStyle.thickThin;
    
    await context.sync();
    
    console.log("Shape border line style set to ThickThin");
});
```

---

### transparency

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the degree of transparency of the line as a value between 0.0 (opaque) and 1.0 (clear).

#### Examples

**Example**: Set a shape's border line transparency to 50% (semi-transparent)

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.shapes;
    const shape = shapes.getFirst();
    
    // Set the line transparency to 0.5 (50% transparent)
    shape.lineFormat.transparency = 0.5;
    
    await context.sync();
});
```

---

### weight

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the thickness of the line in points.

#### Examples

**Example**: Set the thickness of a shape's border line to 3 points

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.shapes;
    const shape = shapes.getItem(0);
    
    // Set the line weight to 3 points
    shape.lineFormat.weight = 3;
    
    await context.sync();
});
```

---

## Methods

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.LineFormatLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.LineFormat`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.LineFormat`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.LineFormat`

#### Examples

**Example**: Load and read the line color and weight properties of the first shape's border in the document

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    const firstShape = shapes.getFirst();
    const lineFormat = firstShape.lineFormat;
    
    // Load specific properties of the line format
    lineFormat.load("color, weight");
    
    // Sync to execute the load command
    await context.sync();
    
    // Now we can read the loaded properties
    console.log("Line color: " + lineFormat.color);
    console.log("Line weight: " + lineFormat.weight);
});
```

---

### set

**Kind:** `configure`

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `Interfaces.LineFormatUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.LineFormat` (required)

  **Returns:** `void`

#### Examples

**Example**: Format a shape's border by setting multiple line properties at once, including color, weight, and dash style.

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    const shape = shapes.items[0];
    const lineFormat = shape.line;
    
    // Set multiple line format properties at once
    lineFormat.set({
        color: "#FF0000",
        weight: 3,
        dashStyle: Word.ShapeLineDashStyle.dash,
        visible: true
    });
    
    await context.sync();
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify().

#### Signature

**Returns:** `Word.Interfaces.LineFormatData`
Whereas the original Word.LineFormat object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.LineFormatData) that contains shallow copies of any loaded child properties from the original object.

#### Examples

**Example**: Get the line format properties of a shape's border as a JSON object and log it to the console for inspection.

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    const shape = shapes.getFirst();
    
    // Get the line format of the shape's border
    const lineFormat = shape.line;
    
    // Load properties to inspect
    lineFormat.load("color,weight,dashStyle,visible");
    
    await context.sync();
    
    // Convert the line format to JSON
    const lineFormatJSON = lineFormat.toJSON();
    
    // Log the JSON representation
    console.log("Line Format Properties:", lineFormatJSON);
    console.log("Color:", lineFormatJSON.color);
    console.log("Weight:", lineFormatJSON.weight);
    console.log("Dash Style:", lineFormatJSON.dashStyle);
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document.

#### Signature

**Returns:** `Word.LineFormat`

#### Examples

**Example**: Format a shape's border with a red color and track the border's line format to automatically adjust when the document changes

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    const shape = shapes.getFirst();
    
    // Get the line format (border) of the shape
    const lineFormat = shape.line;
    
    // Track the line format for automatic adjustment
    lineFormat.track();
    
    // Set border color to red
    lineFormat.color = "red";
    lineFormat.weight = 2;
    
    await context.sync();
    
    console.log("Line format is now tracked and will adjust automatically");
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked.

#### Signature

**Returns:** `Word.LineFormat`

#### Examples

**Example**: Release memory for a line format object after modifying a shape's border properties to optimize memory usage in a long-running add-in.

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    const shape = shapes.getFirst();
    
    // Get the line format and track it
    const lineFormat = shape.line;
    lineFormat.load("color,weight");
    
    await context.sync();
    
    // Modify the line format properties
    lineFormat.color = "blue";
    lineFormat.weight = 3;
    
    await context.sync();
    
    // Release the memory associated with the line format object
    lineFormat.untrack();
    
    console.log("Line format modified and memory released");
});
```

---

## Source

- /en-us/javascript/api/word/word.lineformat
