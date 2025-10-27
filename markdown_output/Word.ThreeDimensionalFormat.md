# ThreeDimensionalFormat

**Package:** `word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents a shape's three-dimensional formatting.

## Properties

### bevelBottomDepth

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the depth of the bottom bevel.

#### Examples

**Example**: Set the bottom bevel depth of a selected shape to 10 points to create a three-dimensional effect.

```typescript
await Word.run(async (context) => {
    const shapes = context.document.getSelection().getShapes();
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        shape.threeDFormat.bevelBottomDepth = 10;
        await context.sync();
    }
});
```

---

### bevelBottomInset

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the inset size for the bottom bevel.

#### Examples

**Example**: Set the bottom bevel inset to 10 points for a selected shape in the document

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        
        // Set the bottom bevel inset to 10 points
        shape.threeDFormat.bevelBottomInset = 10;
        
        await context.sync();
        console.log("Bottom bevel inset set to 10 points");
    }
});
```

---

### bevelBottomType

**Type:** `Word.BevelType | "mixed" | "none" | "relaxedInset" | "circle" | "slope" | "cross" | "angle" | "softRound" | "convex" | "coolSlant" | "divot" | "riblet" | "hardEdge" | "artDeco"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies a BevelType value that represents the bevel type for the bottom bevel.

#### Examples

**Example**: Set the bottom bevel type of a selected shape to "softRound" to give it a smooth, rounded 3D edge effect at the bottom.

```typescript
await Word.run(async (context) => {
    const shapes = context.document.getSelection().inlinePictures;
    shapes.load("items");
    await context.sync();

    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const threeDFormat = shape.threeDimensionalFormat;
        threeDFormat.bevelBottomType = Word.BevelType.softRound;
        
        await context.sync();
        console.log("Bottom bevel type set to softRound");
    }
});
```

---

### bevelTopDepth

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the depth of the top bevel.

#### Examples

**Example**: Set the top bevel depth of a selected shape to 10 points to create a three-dimensional appearance.

```typescript
await Word.run(async (context) => {
    const shapes = context.document.getSelection().getShapes();
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        shape.threeDFormat.bevelTopDepth = 10;
        await context.sync();
    }
});
```

---

### bevelTopInset

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the inset size for the top bevel.

#### Examples

**Example**: Set the top bevel inset size to 10 points for a selected shape in the document

```typescript
await Word.run(async (context) => {
    const shapes = context.document.getSelection().inlinePictures;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const threeDFormat = shape.threeDimensionalFormat;
        threeDFormat.bevelTopInset = 10;
        
        await context.sync();
        console.log("Top bevel inset set to 10 points");
    }
});
```

---

### bevelTopType

**Type:** `Word.BevelType | "mixed" | "none" | "relaxedInset" | "circle" | "slope" | "cross" | "angle" | "softRound" | "convex" | "coolSlant" | "divot" | "riblet" | "hardEdge" | "artDeco"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies a BevelType value that represents the bevel type for the top bevel.

#### Examples

**Example**: Apply a "softRound" bevel type to the top of a 3D shape in the document

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.inlinePictures;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        
        // Set the top bevel type to soft round
        shape.threeDimensionalFormat.bevelTopType = Word.BevelType.softRound;
        
        await context.sync();
        console.log("Top bevel type set to soft round");
    }
});
```

---

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a ThreeDimensionalFormat object to synchronize changes to a shape's 3D properties with the Office host application.

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    const shape = shapes.items[0];
    const threeDFormat = shape.threeDimensionalFormat;
    
    // Access the request context from the ThreeDimensionalFormat object
    const formatContext = threeDFormat.context;
    
    // Use the context to load properties
    threeDFormat.load("visible");
    await formatContext.sync();
    
    console.log("3D Format visible:", threeDFormat.visible);
});
```

---

### contourColor

**Type:** `Word.ColorFormat`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns a ColorFormat object that represents color of the contour of a shape.

#### Examples

**Example**: Set the contour color of a 3D shape to red with full opacity

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.inlinePictures;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const threeDFormat = shape.threeDimensionalFormat;
        const contourColor = threeDFormat.contourColor;
        
        // Set the contour color to red
        contourColor.setSolidColor("#FF0000");
        
        await context.sync();
    }
});
```

---

### contourWidth

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the width of the contour of a shape.

#### Examples

**Example**: Set the contour width of a selected shape to 5 points to create a visible outline effect

```typescript
await Word.run(async (context) => {
    const shapes = context.document.getSelection().getShapes();
    shapes.load("items");
    await context.sync();

    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        shape.threeDFormat.contourWidth = 5;
        await context.sync();
    }
});
```

---

### depth

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the depth of the shape's extrusion.

#### Examples

**Example**: Set the 3D extrusion depth of a selected shape to 50 points to create a deeper three-dimensional effect.

```typescript
await Word.run(async (context) => {
    const shapes = context.document.getSelection().inlinePictures;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const threeDFormat = shape.threeDimensionalFormat;
        threeDFormat.depth = 50;
        await context.sync();
    }
});
```

---

### extrusionColor

**Type:** `Word.ColorFormat`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns a ColorFormat object that represents the color of the shape's extrusion.

#### Examples

**Example**: Set the extrusion color of a selected shape to red

```typescript
await Word.run(async (context) => {
    const shapes = context.document.getSelection().inlinePictures;
    shapes.load("items");
    await context.sync();

    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const threeDFormat = shape.threeDimensionalFormat;
        const extrusionColor = threeDFormat.extrusionColor;
        
        // Set the extrusion color to red
        extrusionColor.setRgb(255, 0, 0);
        
        await context.sync();
    }
});
```

---

### extrusionColorType

**Type:** `Word.ExtrusionColorType | "mixed" | "automatic" | "custom"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether the extrusion color is based on the extruded shape's fill (the front face of the extrusion) and automatically changes when the shape's fill changes, or whether the extrusion color is independent of the shape's fill.

#### Examples

**Example**: Set a shape's extrusion color type to automatic so it changes based on the shape's fill color

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    // Get the first shape
    const shape = shapes.items[0];
    
    // Set the extrusion color type to automatic
    shape.threeDFormat.extrusionColorType = Word.ExtrusionColorType.automatic;
    
    await context.sync();
    console.log("Extrusion color type set to automatic");
});
```

---

### fieldOfView

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the amount of perspective for a shape.

#### Examples

**Example**: Set the field of view to 60 degrees to add perspective to a 3D shape in the document

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    // Get the first shape
    const shape = shapes.items[0];
    const threeDFormat = shape.threeDFormat;
    
    // Set the field of view to 60 degrees for perspective
    threeDFormat.fieldOfView = 60;
    
    await context.sync();
});
```

---

### isPerspective

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies true if the extrusion appears in perspective — that is, if the walls of the extrusion narrow toward a vanishing point, false if the extrusion is a parallel, or orthographic, projection — that is, if the walls don't narrow toward a vanishing point.

#### Examples

**Example**: Enable perspective projection on a shape's 3D extrusion so that the walls narrow toward a vanishing point, creating a realistic depth effect.

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.inlinePictures;
    shapes.load("items");
    await context.sync();

    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const threeDFormat = shape.threeDimensionalFormat;
        
        // Enable perspective projection for realistic 3D depth
        threeDFormat.isPerspective = true;
        
        await context.sync();
        console.log("Perspective projection enabled on shape");
    }
});
```

---

### isVisible

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies if the specified object, or the formatting applied to it, is visible.

#### Examples

**Example**: Make a shape's 3D formatting visible by setting the isVisible property to true

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.inlinePictures;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const threeDFormat = shape.threeDimensionalFormat;
        threeDFormat.isVisible = true;
        
        await context.sync();
        console.log("3D formatting is now visible");
    }
});
```

---

### lightAngle

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the angle of the lighting.

#### Examples

**Example**: Set the 3D lighting angle to 45 degrees for a selected shape in the document

```typescript
await Word.run(async (context) => {
    const shapes = context.document.getSelection().getShapes();
    const shape = shapes.getFirst();
    
    shape.threeDFormat.lightAngle = 45;
    
    await context.sync();
});
```

---

### presetCamera

**Type:** `Word.PresetCamera | "Mixed" | "LegacyObliqueTopLeft" | "LegacyObliqueTop" | "LegacyObliqueTopRight" | "LegacyObliqueLeft" | "LegacyObliqueFront" | "LegacyObliqueRight" | "LegacyObliqueBottomLeft" | "LegacyObliqueBottom" | "LegacyObliqueBottomRight" | "LegacyPerspectiveTopLeft" | "LegacyPerspectiveTop" | "LegacyPerspectiveTopRight" | "LegacyPerspectiveLeft" | "LegacyPerspectiveFront" | "LegacyPerspectiveRight" | "LegacyPerspectiveBottomLeft" | "LegacyPerspectiveBottom" | "LegacyPerspectiveBottomRight" | "OrthographicFront" | "IsometricTopUp" | "IsometricTopDown" | "IsometricBottomUp" | "IsometricBottomDown" | "IsometricLeftUp" | "IsometricLeftDown" | "IsometricRightUp" | "IsometricRightDown" | "IsometricOffAxis1Left" | "IsometricOffAxis1Right" | "IsometricOffAxis1Top" | "IsometricOffAxis2Left" | "IsometricOffAxis2Right" | "IsometricOffAxis2Top" | "IsometricOffAxis3Left" | "IsometricOffAxis3Right" | "IsometricOffAxis3Bottom" | "IsometricOffAxis4Left" | "IsometricOffAxis4Right" | "IsometricOffAxis4Bottom" | "ObliqueTopLeft" | "ObliqueTop" | "ObliqueTopRight" | "ObliqueLeft" | "ObliqueRight" | "ObliqueBottomLeft" | "ObliqueBottom" | "ObliqueBottomRight" | "PerspectiveFront" | "PerspectiveLeft" | "PerspectiveRight" | "PerspectiveAbove" | "PerspectiveBelow" | "PerspectiveAboveLeftFacing" | "PerspectiveAboveRightFacing" | "PerspectiveContrastingLeftFacing" | "PerspectiveContrastingRightFacing" | "PerspectiveHeroicLeftFacing" | "PerspectiveHeroicRightFacing" | "PerspectiveHeroicExtremeLeftFacing" | "PerspectiveHeroicExtremeRightFacing" | "PerspectiveRelaxed" | "PerspectiveRelaxedModerately"`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns a PresetCamera value that represents the camera presets.

#### Examples

**Example**: Set the 3D camera preset to "PerspectiveRelaxed" for the first shape in the document to create a relaxed perspective view.

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.inlinePictures;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const threeDFormat = shape.threeDimensionalFormat;
        threeDFormat.presetCamera = "PerspectiveRelaxed";
        
        await context.sync();
        console.log("3D camera preset applied successfully");
    }
});
```

---

### presetExtrusionDirection

**Type:** `Word.PresetExtrusionDirection | "Mixed" | "BottomRight" | "Bottom" | "BottomLeft" | "Right" | "None" | "Left" | "TopRight" | "Top" | "TopLeft"`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns the direction taken by the extrusion's sweep path leading away from the extruded shape (the front face of the extrusion).

#### Examples

**Example**: Get the extrusion direction of a selected shape and display it to the user, then set it to a top-right direction.

```typescript
await Word.run(async (context) => {
    const shapes = context.document.getSelection().inlinePictures;
    shapes.load("items");
    await context.sync();

    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const threeDFormat = shape.threeDimensionalFormat;
        threeDFormat.load("presetExtrusionDirection");
        await context.sync();

        // Get current extrusion direction
        console.log("Current extrusion direction: " + threeDFormat.presetExtrusionDirection);

        // Set extrusion direction to TopRight
        threeDFormat.presetExtrusionDirection = Word.PresetExtrusionDirection.topRight;
        await context.sync();

        console.log("Extrusion direction set to TopRight");
    }
});
```

---

### presetLighting

**Type:** `Word.LightRigType | "Mixed" | "LegacyFlat1" | "LegacyFlat2" | "LegacyFlat3" | "LegacyFlat4" | "LegacyNormal1" | "LegacyNormal2" | "LegacyNormal3" | "LegacyNormal4" | "LegacyHarsh1" | "LegacyHarsh2" | "LegacyHarsh3" | "LegacyHarsh4" | "ThreePoint" | "Balanced" | "Soft" | "Harsh" | "Flood" | "Contrasting" | "Morning" | "Sunrise" | "Sunset" | "Chilly" | "Freezing" | "Flat" | "TwoPoint" | "Glow" | "BrightRoom"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies a LightRigType value that represents the lighting preset.

#### Examples

**Example**: Apply a "Sunrise" lighting preset to the 3D formatting of the first shape in the document to create a warm, directional lighting effect.

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const threeDFormat = shape.threeDFormat;
        threeDFormat.presetLighting = "Sunrise";
        
        await context.sync();
        console.log("Applied Sunrise lighting preset to the shape");
    }
});
```

---

### presetLightingDirection

**Type:** `Word.PresetLightingDirection | "Mixed" | "TopLeft" | "Top" | "TopRight" | "Left" | "None" | "Right" | "BottomLeft" | "Bottom" | "BottomRight"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the position of the light source relative to the extrusion.

#### Examples

**Example**: Set the 3D lighting direction to come from the top-right for a selected shape in the document.

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        
        // Set the preset lighting direction to top-right
        shape.threeDFormat.presetLightingDirection = Word.PresetLightingDirection.topRight;
        
        await context.sync();
        console.log("Lighting direction set to top-right");
    }
});
```

---

### presetLightingSoftness

**Type:** `Word.PresetLightingSoftness | "Mixed" | "Dim" | "Normal" | "Bright"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the intensity of the extrusion lighting.

#### Examples

**Example**: Set the 3D lighting softness of a selected shape to "Bright" to increase the intensity of the extrusion lighting.

```typescript
await Word.run(async (context) => {
    const shapes = context.document.getSelection().inlinePictures;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const threeDFormat = shape.threeDimensionalFormat;
        threeDFormat.presetLightingSoftness = Word.PresetLightingSoftness.bright;
        
        await context.sync();
        console.log("3D lighting softness set to Bright");
    }
});
```

---

### presetMaterial

**Type:** `Word.PresetMaterial | "Mixed" | "Matte" | "Plastic" | "Metal" | "WireFrame" | "Matte2" | "Plastic2" | "Metal2" | "WarmMatte" | "TranslucentPowder" | "Powder" | "DarkEdge" | "SoftEdge" | "Clear" | "Flat" | "SoftMetal"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the extrusion surface material.

#### Examples

**Example**: Set the 3D surface material of a selected shape to a metallic finish

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        
        // Set the preset material to Metal
        shape.threeDFormat.presetMaterial = Word.PresetMaterial.metal;
        
        await context.sync();
        console.log("Shape material set to Metal");
    }
});
```

---

### presetThreeDimensionalFormat

**Type:** `Word.PresetThreeDimensionalFormat | "Mixed" | "Format1" | "Format2" | "Format3" | "Format4" | "Format5" | "Format6" | "Format7" | "Format8" | "Format9" | "Format10" | "Format11" | "Format12" | "Format13" | "Format14" | "Format15" | "Format16" | "Format17" | "Format18" | "Format19" | "Format20"`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns the preset extrusion format.

#### Examples

**Example**: Get the preset 3D format applied to a shape and display it to the user

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const threeDFormat = shape.threeDimensionalFormat;
        threeDFormat.load("presetThreeDimensionalFormat");
        await context.sync();
        
        console.log("Preset 3D Format: " + threeDFormat.presetThreeDimensionalFormat);
    } else {
        console.log("No shapes found in the document");
    }
});
```

---

### projectText

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether text on a shape rotates with shape. true rotates the text.

#### Examples

**Example**: Enable text rotation on a shape so that the text rotates along with the shape's 3D orientation

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    // Get the first shape
    const shape = shapes.items[0];
    
    // Enable text rotation with the shape
    shape.threeDFormat.projectText = true;
    
    await context.sync();
});
```

---

### rotationX

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the rotation of the extruded shape around the x-axis in degrees.

#### Examples

**Example**: Rotate a shape 45 degrees around the x-axis to create a tilted 3D effect

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.shapes;
    const shape = shapes.getItemAt(0);
    
    // Rotate the shape 45 degrees around the x-axis
    shape.threeDFormat.rotationX = 45;
    
    await context.sync();
});
```

---

### rotationY

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the rotation of the extruded shape around the y-axis in degrees.

#### Examples

**Example**: Rotate a selected shape 45 degrees around the y-axis to create a three-dimensional perspective effect

```typescript
await Word.run(async (context) => {
    const shapes = context.document.getSelection().getShapes();
    shapes.load("items");
    await context.sync();

    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        shape.threeDFormat.rotationY = 45;
        await context.sync();
    }
});
```

---

### rotationZ

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the z-axis rotation of the camera.

#### Examples

**Example**: Set the 3D camera's z-axis rotation to 45 degrees for a selected shape

```typescript
await Word.run(async (context) => {
    const shapes = context.document.getSelection().inlinePictures;
    shapes.load("items");
    
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const threeDFormat = shape.threeDimensionalFormat;
        threeDFormat.rotationZ = 45;
        
        await context.sync();
    }
});
```

---

### z

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the position on the z-axis for the shape.

#### Examples

**Example**: Set a shape's z-axis position to 50 to move it forward in 3D space

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    const shape = shapes.items[0];
    shape.threeDFormat.z = 50;
    
    await context.sync();
});
```

---

## Methods

### incrementRotationHorizontal

**Kind:** `write`

Horizontally rotates a shape on the x-axis. The number of degrees to rotate.

#### Signature

**Parameters:**
- `increment`: `number` (required)

**Returns:** `void`

#### Examples

**Example**: Rotate a selected shape horizontally by 45 degrees on the x-axis

```typescript
await Word.run(async (context) => {
    const shapes = context.document.getSelection().inlinePictures;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const threeDFormat = shape.threeDimensionalFormat;
        threeDFormat.incrementRotationHorizontal(45);
        await context.sync();
    }
});
```

---

### incrementRotationVertical

**Kind:** `write`

Vertically rotates a shape on the y-axis. The number of degrees to rotate.

#### Signature

**Parameters:**
- `increment`: `number` (required)

**Returns:** `void`

#### Examples

**Example**: Rotate a selected shape vertically by 45 degrees on the y-axis to create a 3D tilting effect

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        
        // Get the 3D format of the shape
        const threeDFormat = shape.threeDFormat;
        
        // Rotate the shape vertically by 45 degrees on the y-axis
        threeDFormat.incrementRotationVertical(45);
        
        await context.sync();
        console.log("Shape rotated vertically by 45 degrees");
    }
});
```

---

### incrementRotationX

**Kind:** `write`

Changes the rotation around the x-axis. The number of degrees to rotate.

#### Signature

**Parameters:**
- `increment`: `number` (required)

**Returns:** `void`

#### Examples

**Example**: Rotate a shape 45 degrees around the x-axis to create a tilted 3D effect

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    const shape = shapes.items[0];
    const threeDFormat = shape.threeDimensionalFormat;
    
    // Rotate the shape 45 degrees around the x-axis
    threeDFormat.incrementRotationX(45);
    
    await context.sync();
});
```

---

### incrementRotationY

**Kind:** `write`

Changes the rotation around the y-axis. The number of degrees to rotate.

#### Signature

**Parameters:**
- `increment`: `number` (required)

**Returns:** `void`

#### Examples

**Example**: Rotate a selected shape 45 degrees around the y-axis to create a three-dimensional rotation effect

```typescript
await Word.run(async (context) => {
    const shapes = context.document.getSelection().getShapes();
    shapes.load("items");
    await context.sync();

    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const threeDFormat = shape.threeDFormat;
        
        // Rotate the shape 45 degrees around the y-axis
        threeDFormat.incrementRotationY(45);
        
        await context.sync();
        console.log("Shape rotated 45 degrees around y-axis");
    }
});
```

---

### incrementRotationZ

**Kind:** `write`

Rotates a shape on the z-axis. The number of degrees to rotate.

#### Signature

**Parameters:**
- `increment`: `number` (required)

**Returns:** `void`

#### Examples

**Example**: Rotate a selected shape 45 degrees clockwise around the z-axis (depth rotation)

```typescript
await Word.run(async (context) => {
    const shapes = context.document.getSelection().inlinePictures;
    shapes.load("items");
    await context.sync();

    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const threeDFormat = shape.threeDimensionalFormat;
        
        // Rotate the shape 45 degrees on the z-axis
        threeDFormat.incrementRotationZ(45);
        
        await context.sync();
        console.log("Shape rotated 45 degrees on z-axis");
    }
});
```

---

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.ThreeDimensionalFormatLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.ThreeDimensionalFormat`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.ThreeDimensionalFormat`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.ThreeDimensionalFormat`

#### Examples

**Example**: Load and read the depth property of a shape's 3D formatting to check if it has depth applied.

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    const shape = shapes.items[0];
    const threeDFormat = shape.threeDFormat;
    
    // Load the depth property of the 3D format
    threeDFormat.load("depth");
    await context.sync();
    
    // Read the loaded property
    console.log(`Shape 3D depth: ${threeDFormat.depth}`);
});
```

---

### resetRotation

**Kind:** `write`

Resets the extrusion rotation around the x-axis, y-axis, and z-axis to 0.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Reset the 3D rotation of a selected shape back to its default orientation (0 degrees on all axes)

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        
        // Get the 3D format of the shape
        const threeDFormat = shape.threeDFormat;
        
        // Reset the rotation to 0 on all axes
        threeDFormat.resetRotation();
        
        await context.sync();
        console.log("Shape rotation has been reset to default (0, 0, 0)");
    }
});
```

---

### set

**Kind:** `write`

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `Interfaces.ThreeDimensionalFormatUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.ThreeDimensionalFormat` (required)

  **Returns:** `void`

#### Examples

**Example**: Apply multiple 3D formatting properties to a shape at once, setting depth, contour width, and material type

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    const shape = shapes.items[0];
    const threeDFormat = shape.threeDFormat;
    
    // Set multiple 3D properties at once
    threeDFormat.set({
        depth: 50,
        contourWidth: 5,
        presetMaterial: Word.PresetMaterial.metal
    });
    
    await context.sync();
});
```

---

### setExtrusionDirection

**Kind:** `write`

Sets the direction of the extrusion's sweep path.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `presetExtrusionDirection`: `Word.PresetExtrusionDirection` (required)
    The preset direction.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `presetExtrusionDirection`: `"Mixed" | "BottomRight" | "Bottom" | "BottomLeft" | "Right" | "None" | "Left" | "TopRight" | "Top" | "TopLeft"` (required)
    The preset direction.

  **Returns:** `void`

#### Examples

**Example**: Apply a 3D extrusion effect to a shape with a bottom-right direction to create a perspective effect

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    const shape = shapes.items[0];
    
    // Get the 3D format of the shape
    const threeDFormat = shape.threeDFormat;
    
    // Set the extrusion direction to bottom-right
    threeDFormat.setExtrusionDirection(Word.PresetExtrusionDirection.bottomRight);
    
    await context.sync();
    
    console.log("Extrusion direction set to bottom-right");
});
```

---

### setPresetCamera

**Kind:** `write`

Sets the camera preset for the shape. The preset camera type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `presetCamera`: `Word.PresetCamera` (required)

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `presetCamera`: `"Mixed" | "LegacyObliqueTopLeft" | "LegacyObliqueTop" | "LegacyObliqueTopRight" | "LegacyObliqueLeft" | "LegacyObliqueFront" | "LegacyObliqueRight" | "LegacyObliqueBottomLeft" | "LegacyObliqueBottom" | "LegacyObliqueBottomRight" | "LegacyPerspectiveTopLeft" | "LegacyPerspectiveTop" | "LegacyPerspectiveTopRight" | "LegacyPerspectiveLeft" | "LegacyPerspectiveFront" | "LegacyPerspectiveRight" | "LegacyPerspectiveBottomLeft" | "LegacyPerspectiveBottom" | "LegacyPerspectiveBottomRight" | "OrthographicFront" | "IsometricTopUp" | "IsometricTopDown" | "IsometricBottomUp" | "IsometricBottomDown" | "IsometricLeftUp" | "IsometricLeftDown" | "IsometricRightUp" | "IsometricRightDown" | "IsometricOffAxis1Left" | "IsometricOffAxis1Right" | "IsometricOffAxis1Top" | "IsometricOffAxis2Left" | "IsometricOffAxis2Right" | "IsometricOffAxis2Top" | "IsometricOffAxis3Left" | "IsometricOffAxis3Right" | "IsometricOffAxis3Bottom" | "IsometricOffAxis4Left" | "IsometricOffAxis4Right" | "IsometricOffAxis4Bottom" | "ObliqueTopLeft" | "ObliqueTop" | "ObliqueTopRight" | "ObliqueLeft" | "ObliqueRight" | "ObliqueBottomLeft" | "ObliqueBottom" | "ObliqueBottomRight" | "PerspectiveFront" | "PerspectiveLeft" | "PerspectiveRight" | "PerspectiveAbove" | "PerspectiveBelow" | "PerspectiveAboveLeftFacing" | "PerspectiveAboveRightFacing" | "PerspectiveContrastingLeftFacing" | "PerspectiveContrastingRightFacing" | "PerspectiveHeroicLeftFacing" | "PerspectiveHeroicRightFacing" | "PerspectiveHeroicExtremeLeftFacing" | "PerspectiveHeroicExtremeRightFacing" | "PerspectiveRelaxed" | "PerspectiveRelaxedModerately"` (required)

  **Returns:** `void`

#### Examples

**Example**: Apply a 3D isometric left camera preset to the first shape in the document to create a three-dimensional viewing perspective.

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const threeDFormat = shape.threeDFormat;
        threeDFormat.setPresetCamera(Word.PresetCameraType.isometricLeft);
        await context.sync();
    }
});
```

---

### setThreeDimensionalFormat

**Kind:** `write`

Sets the preset extrusion format. The preset format.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `presetThreeDimensionalFormat`: `Word.PresetThreeDimensionalFormat` (required)

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `presetThreeDimensionalFormat`: `"Mixed" | "Format1" | "Format2" | "Format3" | "Format4" | "Format5" | "Format6" | "Format7" | "Format8" | "Format9" | "Format10" | "Format11" | "Format12" | "Format13" | "Format14" | "Format15" | "Format16" | "Format17" | "Format18" | "Format19" | "Format20"` (required)

  **Returns:** `void`

#### Examples

**Example**: Apply a preset 3D extrusion format to a shape in the document

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        
        // Get the 3D format of the shape and set a preset 3D format
        const threeDFormat = shape.threeDimensionalFormat;
        threeDFormat.setThreeDimensionalFormat(Word.PresetThreeDimensionalFormat.preset1);
        
        await context.sync();
        console.log("Preset 3D format applied to the shape");
    }
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.ThreeDimensionalFormat object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.ThreeDimensionalFormatData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.ThreeDimensionalFormatData`

#### Examples

**Example**: Get the 3D formatting properties of a shape as a plain JavaScript object and log it to the console for inspection or serialization.

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const threeDFormat = shape.threeDFormat;
        
        // Load the 3D format properties
        threeDFormat.load("*");
        await context.sync();
        
        // Convert to plain JavaScript object
        const threeDFormatData = threeDFormat.toJSON();
        
        // Log the plain object (useful for debugging or serialization)
        console.log("3D Format Properties:", JSON.stringify(threeDFormatData, null, 2));
    }
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.ThreeDimensionalFormat`

#### Examples

**Example**: Apply 3D formatting to a shape and track it across multiple sync calls to maintain the reference while modifying its properties

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.inlinePictures;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const threeDFormat = shape.threeDimensionalFormat;
        
        // Track the 3D format object to use it across sync calls
        threeDFormat.track();
        
        // First sync: load current properties
        threeDFormat.load("depth");
        await context.sync();
        
        console.log("Current depth: " + threeDFormat.depth);
        
        // Second sync: modify properties
        threeDFormat.depth = 50;
        await context.sync();
        
        // Untrack when done
        threeDFormat.untrack();
    }
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.ThreeDimensionalFormat`

#### Examples

**Example**: Apply 3D formatting to a shape, then untrack the ThreeDimensionalFormat object to free memory after the formatting is complete.

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        
        // Get and track the 3D format object
        const threeDFormat = shape.threeDFormat;
        context.trackedObjects.add(threeDFormat);
        
        // Apply 3D formatting properties
        threeDFormat.load("presetCamera");
        await context.sync();
        
        // Modify 3D properties
        threeDFormat.presetCamera = Word.PresetCameraType.isometricLeftDown;
        await context.sync();
        
        // Untrack the object to release memory
        threeDFormat.untrack();
        await context.sync();
        
        console.log("3D formatting applied and object untracked");
    }
});
```

---

## Source

- https://learn.microsoft.com/en-us/javascript/api/word/word.threedimensionalformat
