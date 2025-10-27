# Word.FillFormat

**Package:** `word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents the fill formatting for a shape or text.

## Properties

### backgroundColor

**Type:** `Word.ColorFormat`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns a ColorFormat object that represents the background color for the fill.

#### Examples

**Example**: Set the background color of a shape's fill to light blue

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    const shape = shapes.items[0];
    
    // Access the fill format and set background color to light blue
    const fillFormat = shape.fill;
    fillFormat.backgroundColor.set("#ADD8E6");
    
    await context.sync();
});
```

---

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a shape's fill format to verify the connection to the Word host application and log context information.

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const fillFormat = shape.fill;
        
        // Access the request context from the fill format
        const fillContext = fillFormat.context;
        
        // Verify the context is connected to the same Word context
        console.log("Context is valid:", fillContext === context);
        console.log("Context type:", fillContext.constructor.name);
        
        // Use the context to perform operations
        fillFormat.load("type");
        await fillContext.sync();
        
        console.log("Fill type:", fillFormat.type);
    }
});
```

---

### foregroundColor

**Type:** `Word.ColorFormat`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns a ColorFormat object that represents the foreground color for the fill.

#### Examples

**Example**: Set the foreground color of a shape's fill to red

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const fill = shape.fill;
        fill.foregroundColor.set("#FF0000");
        
        await context.sync();
    }
});
```

---

### gradientAngle

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the angle of the gradient fill. The valid range of values is from 0 to 359.9.

#### Examples

**Example**: Set a shape's gradient fill angle to 45 degrees to create a diagonal gradient effect.

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.getShapes();
    const shape = shapes.getFirst();
    
    // Access the fill format and set gradient angle to 45 degrees
    shape.fill.gradientAngle = 45;
    
    await context.sync();
});
```

---

### gradientColorType

**Type:** `Word.GradientColorType | "Mixed" | "OneColor" | "TwoColors" | "PresetColors" | "MultiColor"`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the gradient color type.

#### Examples

**Example**: Check if a shape's fill has a gradient and display the gradient color type (e.g., OneColor, TwoColors, PresetColors).

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.inlinePictures;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const fill = shape.fill;
        fill.load("gradientColorType");
        await context.sync();
        
        console.log("Gradient color type: " + fill.gradientColorType);
        // Possible values: "Mixed", "OneColor", "TwoColors", "PresetColors", "MultiColor"
    }
});
```

---

### gradientDegree

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns how dark or light a one-color gradient fill is. A value of 0 means that black is mixed in with the shape's foreground color to form the gradient. A value of 1 means that white is mixed in. Values between 0 and 1 mean that a darker or lighter shade of the foreground color is mixed in.

#### Examples

**Example**: Check if a shape has a one-color gradient fill and display how dark or light it is by reading the gradient degree value.

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.inlinePictures;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const fill = shape.fill;
        fill.load("gradientDegree");
        await context.sync();
        
        const degree = fill.gradientDegree;
        console.log(`Gradient degree: ${degree}`);
        // 0 = black mixed in, 1 = white mixed in, 0-1 = shade variation
    }
});
```

---

### gradientStyle

**Type:** `Word.GradientStyle | "Mixed" | "Horizontal" | "Vertical" | "DiagonalUp" | "DiagonalDown" | "FromCorner" | "FromTitle" | "FromCenter"`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns the gradient style for the fill.

#### Examples

**Example**: Get the gradient style of a shape's fill and display it to the user

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const fill = shape.fill;
        
        // Load the gradient style property
        fill.load("gradientStyle");
        await context.sync();
        
        // Display the gradient style
        console.log(`Gradient style: ${fill.gradientStyle}`);
        // Possible values: "Horizontal", "Vertical", "DiagonalUp", "DiagonalDown", 
        // "FromCorner", "FromTitle", "FromCenter", or "Mixed"
    }
});
```

---

### gradientVariant

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns the gradient variant for the fill as an integer value from 1 to 4 for most gradient fills.

#### Examples

**Example**: Check if a shape's gradient fill is using the first variant and display the variant number in the console.

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.inlinePictures;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const fill = shape.fill;
        fill.load("type,gradientVariant");
        await context.sync();
        
        if (fill.type === Word.FillType.gradient) {
            console.log(`Gradient variant: ${fill.gradientVariant}`);
            
            if (fill.gradientVariant === 1) {
                console.log("This shape is using gradient variant 1");
            }
        }
    }
});
```

---

### isVisible

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies if the object, or the formatting applied to it, is visible.

#### Examples

**Example**: Hide the fill formatting of a shape to make it transparent

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        
        // Hide the fill formatting to make the shape transparent
        shape.fill.isVisible = false;
        
        await context.sync();
    }
});
```

---

### pattern

**Type:** `Word.PatternType | "Mixed" | "Percent5" | "Percent10" | "Percent20" | "Percent25" | "Percent30" | "Percent40" | "Percent50" | "Percent60" | "Percent70" | "Percent75" | "Percent80" | "Percent90" | "DarkHorizontal" | "DarkVertical" | "DarkDownwardDiagonal" | "DarkUpwardDiagonal" | "SmallCheckerBoard" | "Trellis" | "LightHorizontal" | "LightVertical" | "LightDownwardDiagonal" | "LightUpwardDiagonal" | "SmallGrid" | "DottedDiamond" | "WideDownwardDiagonal" | "WideUpwardDiagonal" | "DashedUpwardDiagonal" | "DashedDownwardDiagonal" | "NarrowVertical" | "NarrowHorizontal" | "DashedVertical" | "DashedHorizontal" | "LargeConfetti" | "LargeGrid" | "HorizontalBrick" | "LargeCheckerBoard" | "SmallConfetti" | "ZigZag" | "SolidDiamond" | "DiagonalBrick" | "OutlinedDiamond" | "Plaid" | "Sphere" | "Weave" | "DottedGrid" | "Divot" | "Shingle" | "Wave" | "Horizontal" | "Vertical" | "Cross" | "DownwardDiagonal" | "UpwardDiagonal" | "DiagonalCross"`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns a PatternType value that represents the pattern applied to the fill or line.

#### Examples

**Example**: Get the pattern type of a shape's fill and display it, then set the fill to use a diagonal cross pattern.

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.inlinePictures;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const fill = shape.fill;
        
        // Get the current pattern
        fill.load("pattern");
        await context.sync();
        
        console.log("Current pattern: " + fill.pattern);
        
        // Set the fill to use a diagonal cross pattern
        fill.pattern = Word.PatternType.diagonalCross;
        
        await context.sync();
    }
});
```

---

### presetGradientType

**Type:** `Word.PresetGradientType | "Mixed" | "EarlySunset" | "LateSunset" | "Nightfall" | "Daybreak" | "Horizon" | "Desert" | "Ocean" | "CalmWater" | "Fire" | "Fog" | "Moss" | "Peacock" | "Wheat" | "Parchment" | "Mahogany" | "Rainbow" | "RainbowII" | "Gold" | "GoldII" | "Brass" | "Chrome" | "ChromeII" | "Silver" | "Sapphire"`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns the preset gradient type for the fill.

#### Examples

**Example**: Get the preset gradient type of a shape's fill and display it to the user

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const fill = shape.fill;
        fill.load("presetGradientType");
        await context.sync();
        
        console.log("Preset gradient type: " + fill.presetGradientType);
        // Example output: "Preset gradient type: Ocean" or "Mixed"
    }
});
```

---

### presetTexture

**Type:** `Word.PresetTexture | "Mixed" | "Papyrus" | "Canvas" | "Denim" | "WovenMat" | "WaterDroplets" | "PaperBag" | "FishFossil" | "Sand" | "GreenMarble" | "WhiteMarble" | "BrownMarble" | "Granite" | "Newsprint" | "RecycledPaper" | "Parchment" | "Stationery" | "BlueTissuePaper" | "PinkTissuePaper" | "PurpleMesh" | "Bouquet" | "Cork" | "Walnut" | "Oak" | "MediumWood"`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the preset texture.

#### Examples

**Example**: Get the preset texture of a shape's fill format and display it in the console.

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const fillFormat = shape.fill;
        fillFormat.load("presetTexture");
        await context.sync();
        
        console.log("Shape preset texture: " + fillFormat.presetTexture);
    }
});
```

---

### rotateWithObject

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether the fill rotates with the shape.

#### Examples

**Example**: Configure a shape's fill to rotate along with the shape when it is rotated

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.shapes;
    const shape = shapes.getItemAt(0);
    
    // Set the fill to rotate with the shape
    shape.fill.rotateWithObject = true;
    
    await context.sync();
    console.log("Fill is now set to rotate with the shape");
});
```

---

### textureAlignment

**Type:** `Word.TextureAlignment | "Mixed" | "TopLeft" | "Top" | "TopRight" | "Left" | "Center" | "Right" | "BottomLeft" | "Bottom" | "BottomRight"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the alignment (the origin of the coordinate grid) for the tiling of the texture fill.

#### Examples

**Example**: Set the texture alignment of a shape's fill to center position

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        
        // Set the texture alignment to center
        shape.fill.textureAlignment = Word.TextureAlignment.center;
        
        await context.sync();
        console.log("Texture alignment set to center");
    }
});
```

---

### textureHorizontalScale

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the horizontal scaling factor for the texture fill.

#### Examples

**Example**: Set the horizontal scaling factor of a shape's texture fill to 150% to stretch the texture pattern horizontally

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.shapes;
    const shape = shapes.getItemAt(0);
    
    // Access the fill format and set horizontal texture scale to 150%
    shape.fill.textureHorizontalScale = 1.5;
    
    await context.sync();
});
```

---

### textureName

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns the name of the custom texture file for the fill.

#### Examples

**Example**: Get the custom texture file name from a shape's fill and display it in the console.

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const fill = shape.fill;
        
        // Load the texture name property
        fill.load("textureName");
        await context.sync();
        
        // Display the custom texture file name
        console.log("Texture name: " + fill.textureName);
    }
});
```

---

### textureOffsetX

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the horizontal offset of the texture from the origin in points.

#### Examples

**Example**: Set the horizontal texture offset to 25 points for a shape's fill to adjust the texture pattern position

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.getShapes();
    const shape = shapes.getFirst();
    
    // Set the horizontal texture offset to 25 points
    shape.fill.textureOffsetX = 25;
    
    await context.sync();
});
```

---

### textureOffsetY

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the vertical offset of the texture.

#### Examples

**Example**: Set the vertical texture offset to 25 pixels for a shape's fill to adjust the texture pattern position

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.shapes;
    const shape = shapes.getFirst();
    
    // Set the vertical offset of the texture fill
    shape.fill.textureOffsetY = 25;
    
    await context.sync();
});
```

---

### textureTile

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether the texture is tiled.

#### Examples

**Example**: Enable texture tiling for a shape's fill format to create a repeating pattern effect

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        
        // Enable texture tiling for the shape's fill
        shape.fill.textureTile = true;
        
        await context.sync();
        console.log("Texture tiling enabled for the shape");
    }
});
```

---

### textureType

**Type:** `Word.TextureType | "Mixed" | "Preset" | "UserDefined"`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns the texture type for the fill.

#### Examples

**Example**: Check if a shape's fill uses a texture and display the texture type in the console.

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.inlinePictures;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const fill = shape.fill;
        
        // Load the texture type property
        fill.load("textureType");
        await context.sync();
        
        // Display the texture type
        console.log(`Fill texture type: ${fill.textureType}`);
        
        // Check if it's a specific texture type
        if (fill.textureType === Word.TextureType.preset) {
            console.log("The fill uses a preset texture pattern");
        } else if (fill.textureType === Word.TextureType.userDefined) {
            console.log("The fill uses a user-defined texture");
        }
    }
});
```

---

### textureVerticalScale

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the vertical scaling factor for the texture fill as a value between 0.0 and 1.0.

#### Examples

**Example**: Set the vertical scaling factor of a shape's texture fill to 0.5 to compress the texture pattern vertically to half its original height

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.inlinePictures;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const fill = shape.fill;
        
        // Set the texture vertical scale to 0.5 (50% of original height)
        fill.textureVerticalScale = 0.5;
        
        await context.sync();
        console.log("Texture vertical scale set to 0.5");
    }
});
```

---

### transparency

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the degree of transparency of the fill for a shape as a value between 0.0 (opaque) and 1.0 (clear).

#### Examples

**Example**: Set a rectangle shape's fill to 50% transparent (semi-transparent blue)

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.shapes;
    const rectangle = shapes.addShape("Rectangle", 100, 100, 200, 100);
    
    // Set the fill color to blue
    rectangle.fill.setSolidColor("#0000FF");
    
    // Set transparency to 0.5 (50% transparent)
    rectangle.fill.transparency = 0.5;
    
    await context.sync();
});
```

---

### type

**Type:** `Word.FillType | "Mixed" | "Solid" | "Patterned" | "Gradient" | "Textured" | "Background" | "Picture"`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the fill format type.

#### Examples

**Example**: Check the fill type of a shape and display different messages based on whether it has a solid fill, gradient fill, or other fill type.

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.inlinePictures;
    shapes.load("items");
    await context.sync();

    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const fill = shape.fill;
        fill.load("type");
        await context.sync();

        switch (fill.type) {
            case Word.FillType.solid:
            case "Solid":
                console.log("The shape has a solid color fill.");
                break;
            case Word.FillType.gradient:
            case "Gradient":
                console.log("The shape has a gradient fill.");
                break;
            case Word.FillType.picture:
            case "Picture":
                console.log("The shape has a picture fill.");
                break;
            default:
                console.log(`The shape has a ${fill.type} fill type.`);
                break;
        }
    }
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
  - `options`: `Word.Interfaces.FillFormatLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.FillFormat`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.FillFormat`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.FillFormat`

#### Examples

**Example**: Get and display the fill color of the first shape in the document by loading its fill format properties.

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    const firstShape = shapes.getFirst();
    const fillFormat = firstShape.fill;
    
    // Load the fill format properties
    fillFormat.load("type, foregroundColor, transparency");
    
    await context.sync();
    
    // Display the loaded properties
    console.log("Fill type: " + fillFormat.type);
    console.log("Fill color: " + fillFormat.foregroundColor);
    console.log("Transparency: " + fillFormat.transparency);
});
```

---

### set

**Kind:** `write`

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `Interfaces.FillFormatUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.FillFormat` (required)

  **Returns:** `void`

#### Examples

**Example**: Set the fill color and transparency of a shape to create a semi-transparent blue background

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    const shape = shapes.items[0];
    const fill = shape.fill;
    
    // Set multiple fill properties at once
    fill.set({
        foregroundColor: "#4472C4",
        transparency: 0.5,
        visible: true
    });
    
    await context.sync();
});
```

---

### setOneColorGradient

**Kind:** `write`

Sets the fill to a one-color gradient.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `style`: `Word.GradientStyle` (required)
    The gradient style.
  - `variant`: `number` (required)
    The gradient variant. Can be a value from 1 to 4.
  - `degree`: `number` (required)
    The gradient degree. Can be a value from 0.0 (dark) to 1.0 (light).

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `style`: `"Mixed" | "Horizontal" | "Vertical" | "DiagonalUp" | "DiagonalDown" | "FromCorner" | "FromTitle" | "FromCenter"` (required)
    The gradient style.
  - `variant`: `number` (required)
    The gradient variant. Can be a value from 1 to 4.
  - `degree`: `number` (required)
    The gradient degree. Can be a value from 0.0 (dark) to 1.0 (light).

  **Returns:** `void`

#### Examples

**Example**: Apply a one-color gradient fill with a horizontal style to the first shape in the document

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        
        // Set a one-color gradient fill
        // style: horizontal gradient (1)
        // variant: 1 (first variant)
        // degree: 0.5 (50% lightness)
        shape.fill.setOneColorGradient(1, 1, 0.5);
        
        await context.sync();
    }
});
```

---

### setPatterned

**Kind:** `write`

Sets the fill to a pattern.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `pattern`: `Word.PatternType` (required)

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `pattern`: `"Mixed" | "Percent5" | "Percent10" | "Percent20" | "Percent25" | "Percent30" | "Percent40" | "Percent50" | "Percent60" | "Percent70" | "Percent75" | "Percent80" | "Percent90" | "DarkHorizontal" | "DarkVertical" | "DarkDownwardDiagonal" | "DarkUpwardDiagonal" | "SmallCheckerBoard" | "Trellis" | "LightHorizontal" | "LightVertical" | "LightDownwardDiagonal" | "LightUpwardDiagonal" | "SmallGrid" | "DottedDiamond" | "WideDownwardDiagonal" | "WideUpwardDiagonal" | "DashedUpwardDiagonal" | "DashedDownwardDiagonal" | "NarrowVertical" | "NarrowHorizontal" | "DashedVertical" | "DashedHorizontal" | "LargeConfetti" | "LargeGrid" | "HorizontalBrick" | "LargeCheckerBoard" | "SmallConfetti" | "ZigZag" | "SolidDiamond" | "DiagonalBrick" | "OutlinedDiamond" | "Plaid" | "Sphere" | "Weave" | "DottedGrid" | "Divot" | "Shingle" | "Wave" | "Horizontal" | "Vertical" | "Cross" | "DownwardDiagonal" | "UpwardDiagonal" | "DiagonalCross"` (required)

  **Returns:** `void`

#### Examples

**Example**: Set the fill of the first shape in the document to a diagonal stripe pattern

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const fill = shape.fill;
        fill.setPatterned(Word.ShapePatternType.diagonalStripe);
        await context.sync();
    }
});
```

---

### setPresetGradient

**Kind:** `write`

Sets the fill to a preset gradient. The gradient style. The gradient variant. Can be a value from 1 to 4. The preset gradient type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `style`: `Word.GradientStyle` (required)
  - `variant`: `number` (required)
  - `presetGradientType`: `Word.PresetGradientType` (required)

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `style`: `"Mixed" | "Horizontal" | "Vertical" | "DiagonalUp" | "DiagonalDown" | "FromCorner" | "FromTitle" | "FromCenter"` (required)
  - `variant`: `number` (required)
  - `presetGradientType`: `"Mixed" | "EarlySunset" | "LateSunset" | "Nightfall" | "Daybreak" | "Horizon" | "Desert" | "Ocean" | "CalmWater" | "Fire" | "Fog" | "Moss" | "Peacock" | "Wheat" | "Parchment" | "Mahogany" | "Rainbow" | "RainbowII" | "Gold" | "GoldII" | "Brass" | "Chrome" | "ChromeII" | "Silver" | "Sapphire"` (required)

  **Returns:** `void`

#### Examples

**Example**: Apply a preset gradient fill with a horizontal style to the first shape in the document

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const fill = shape.fill;
        
        // Set preset gradient: horizontal style, variant 1, early sunset gradient
        fill.setPresetGradient(
            Word.ShapeLineGradientStyle.horizontal,
            1,
            Word.PresetGradientType.earlySunset
        );
        
        await context.sync();
    }
});
```

---

### setPresetTextured

**Kind:** `write`

Sets the fill to a preset texture.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `presetTexture`: `Word.PresetTexture` (required)

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `presetTexture`: `"Mixed" | "Papyrus" | "Canvas" | "Denim" | "WovenMat" | "WaterDroplets" | "PaperBag" | "FishFossil" | "Sand" | "GreenMarble" | "WhiteMarble" | "BrownMarble" | "Granite" | "Newsprint" | "RecycledPaper" | "Parchment" | "Stationery" | "BlueTissuePaper" | "PinkTissuePaper" | "PurpleMesh" | "Bouquet" | "Cork" | "Walnut" | "Oak" | "MediumWood"` (required)

  **Returns:** `void`

#### Examples

**Example**: Apply a preset textured fill (papyrus texture) to the first shape in the document

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        
        // Set the fill to a preset papyrus texture
        shape.fill.setPresetTextured(Word.PresetTexture.papyrus);
        
        await context.sync();
    }
});
```

---

### setTwoColorGradient

**Kind:** `write`

Sets the fill to a two-color gradient.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `style`: `Word.GradientStyle` (required)
  - `variant`: `number` (required)

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `style`: `"Mixed" | "Horizontal" | "Vertical" | "DiagonalUp" | "DiagonalDown" | "FromCorner" | "FromTitle" | "FromCenter"` (required)
  - `variant`: `number` (required)

  **Returns:** `void`

#### Examples

**Example**: Apply a two-color gradient fill to a shape with a horizontal linear style

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    const shape = shapes.items[0];
    
    // Apply a two-color gradient fill with horizontal linear style
    shape.fill.setTwoColorGradient(
        Word.GradientStyle.horizontal,
        1 // variant number (1-4 depending on style)
    );
    
    await context.sync();
});
```

---

### solid

**Kind:** `write`

Sets the fill to a uniform color.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Set the fill color of the first shape in the document to a solid red color

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const fill = shape.fill;
        fill.solid("#FF0000");
        await context.sync();
    }
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.FillFormat object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.FillFormatData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.FillFormatData`

#### Examples

**Example**: Get the fill format properties of the first shape in the document as a plain JavaScript object and log it to the console.

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.inlinePictures;
    const shape = shapes.getFirstOrNullObject();
    
    // Load the fill format properties
    shape.load("imageFormat/fill");
    await context.sync();
    
    if (!shape.isNullObject) {
        const fillFormat = shape.imageFormat.fill;
        
        // Convert the FillFormat object to a plain JavaScript object
        const fillData = fillFormat.toJSON();
        
        // Log the plain object (useful for debugging or data export)
        console.log("Fill Format Data:", fillData);
        console.log("Fill Type:", fillData.type);
        console.log("Fill Transparency:", fillData.transparency);
    } else {
        console.log("No shapes found in the document.");
    }
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.FillFormat`

#### Examples

**Example**: Track a shape's fill format object to maintain its reference across multiple sync calls while changing its color properties in separate batches.

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.inlinePictures;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const fillFormat = shape.fill;
        
        // Track the fill format object to use it across sync calls
        fillFormat.track();
        
        // First batch: Set fill to solid color
        fillFormat.setSolidColor("#FF6B6B");
        await context.sync();
        
        // Second batch: Change the color (object reference still valid due to tracking)
        fillFormat.setSolidColor("#4ECDC4");
        await context.sync();
        
        // Untrack when done to free up memory
        fillFormat.untrack();
    }
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.FillFormat`

#### Examples

**Example**: Get a shape's fill format, modify its properties, and then untrack it to release memory after the operation is complete.

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const fillFormat = shape.fill;
        
        // Track the fill format object to work with it
        fillFormat.load("type");
        await context.sync();
        
        // Set fill color
        fillFormat.setSolidColor("#FF6347");
        await context.sync();
        
        // Untrack the fill format object to release memory
        fillFormat.untrack();
        await context.sync();
        
        console.log("Fill format modified and untracked successfully");
    }
});
```

---

## Source

- /en-us/javascript/api/word/word.fillformat
