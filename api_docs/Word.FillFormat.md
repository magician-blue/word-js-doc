# Word.FillFormat class

Package: [word](/en-us/javascript/api/word)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents the fill formatting for a shape or text.

Extends: [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks

[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- [backgroundColor](#backgroundcolor)
  - Returns a ColorFormat object that represents the background color for the fill.
- [context](#context)
  - The request context associated with the object. This connects the add-in's process to the Office host application's process.
- [foregroundColor](#foregroundcolor)
  - Returns a ColorFormat object that represents the foreground color for the fill.
- [gradientAngle](#gradientangle)
  - Specifies the angle of the gradient fill. The valid range of values is from 0 to 359.9.
- [gradientColorType](#gradientcolortype)
  - Gets the gradient color type.
- [gradientDegree](#gradientdegree)
  - Returns how dark or light a one-color gradient fill is. A value of 0 means that black is mixed in with the shape's foreground color to form the gradient. A value of 1 means that white is mixed in. Values between 0 and 1 mean that a darker or lighter shade of the foreground color is mixed in.
- [gradientStyle](#gradientstyle)
  - Returns the gradient style for the fill.
- [gradientVariant](#gradientvariant)
  - Returns the gradient variant for the fill as an integer value from 1 to 4 for most gradient fills.
- [isVisible](#isvisible)
  - Specifies if the object, or the formatting applied to it, is visible.
- [pattern](#pattern)
  - Returns a PatternType value that represents the pattern applied to the fill or line.
- [presetGradientType](#presetgradienttype)
  - Returns the preset gradient type for the fill.
- [presetTexture](#presettexture)
  - Gets the preset texture.
- [rotateWithObject](#rotatewithobject)
  - Specifies whether the fill rotates with the shape.
- [textureAlignment](#texturealignment)
  - Specifies the alignment (the origin of the coordinate grid) for the tiling of the texture fill.
- [textureHorizontalScale](#texturehorizontalscale)
  - Specifies the horizontal scaling factor for the texture fill.
- [textureName](#texturename)
  - Returns the name of the custom texture file for the fill.
- [textureOffsetX](#textureoffsetx)
  - Specifies the horizontal offset of the texture from the origin in points.
- [textureOffsetY](#textureoffsety)
  - Specifies the vertical offset of the texture.
- [textureTile](#texturetile)
  - Specifies whether the texture is tiled.
- [textureType](#texturetype)
  - Returns the texture type for the fill.
- [textureVerticalScale](#textureverticalscale)
  - Specifies the vertical scaling factor for the texture fill as a value between 0.0 and 1.0.
- [transparency](#transparency)
  - Specifies the degree of transparency of the fill for a shape as a value between 0.0 (opaque) and 1.0 (clear).
- [type](#type)
  - Gets the fill format type.

## Methods

- [load(options)](#loadoptions)
  - Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- [load(propertyNames)](#loadpropertynames)
  - Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- [load(propertyNamesAndPaths)](#loadpropertynamesandpaths)
  - Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- [set(properties, options)](#setproperties-options)
  - Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
- [set(properties)](#setproperties)
  - Sets multiple properties on the object at the same time, based on an existing loaded object.
- [setOneColorGradient(style, variant, degree)](#setonecolorgradientstyle-variant-degree-1)
  - Sets the fill to a one-color gradient.
- [setOneColorGradient(style, variant, degree)](#setonecolorgradientstyle-variant-degree-2)
  - Sets the fill to a one-color gradient.
- [setPatterned(pattern)](#setpatternedpattern-1)
  - Sets the fill to a pattern.
- [setPatterned(pattern)](#setpatternedpattern-2)
  - Sets the fill to a pattern.
- [setPresetGradient(style, variant, presetGradientType)](#setpresetgradientstyle-variant-presetgradienttype-1)
  - Sets the fill to a preset gradient. The gradient style. The gradient variant. Can be a value from 1 to 4. The preset gradient type.
- [setPresetGradient(style, variant, presetGradientType)](#setpresetgradientstyle-variant-presetgradienttype-2)
  - Sets the fill to a preset gradient. The gradient style. The gradient variant. Can be a value from 1 to 4. The preset gradient type.
- [setPresetTextured(presetTexture)](#setpresettexturedpresettexture-1)
  - Sets the fill to a preset texture.
- [setPresetTextured(presetTexture)](#setpresettexturedpresettexture-2)
  - Sets the fill to a preset texture.
- [setTwoColorGradient(style, variant)](#settwocolorgradientstyle-variant-1)
  - Sets the fill to a two-color gradient.
- [setTwoColorGradient(style, variant)](#settwocolorgradientstyle-variant-2)
  - Sets the fill to a two-color gradient.
- [solid()](#solid)
  - Sets the fill to a uniform color.
- [toJSON()](#tojson)
  - Overrides the JavaScript toJSON() method to provide more useful output for JSON.stringify().
- [track()](#track)
  - Track the object for automatic adjustment based on surrounding changes in the document.
- [untrack()](#untrack)
  - Release the memory associated with this object, if it has previously been tracked.

## Property Details

### backgroundColor

Returns a ColorFormat object that represents the background color for the fill.

```typescript
readonly backgroundColor: Word.ColorFormat;
```

Property Value:
- [Word.ColorFormat](/en-us/javascript/api/word/word.colorformat)

Remarks:
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### context

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value:
- [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### foregroundColor

Returns a ColorFormat object that represents the foreground color for the fill.

```typescript
readonly foregroundColor: Word.ColorFormat;
```

Property Value:
- [Word.ColorFormat](/en-us/javascript/api/word/word.colorformat)

Remarks:
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### gradientAngle

Specifies the angle of the gradient fill. The valid range of values is from 0 to 359.9.

```typescript
gradientAngle: number;
```

Property Value:
- number

Remarks:
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### gradientColorType

Gets the gradient color type.

```typescript
readonly gradientColorType: Word.GradientColorType | "Mixed" | "OneColor" | "TwoColors" | "PresetColors" | "MultiColor";
```

Property Value:
- [Word.GradientColorType](/en-us/javascript/api/word/word.gradientcolortype) | "Mixed" | "OneColor" | "TwoColors" | "PresetColors" | "MultiColor"

Remarks:
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### gradientDegree

Returns how dark or light a one-color gradient fill is. A value of 0 means that black is mixed in with the shape's foreground color to form the gradient. A value of 1 means that white is mixed in. Values between 0 and 1 mean that a darker or lighter shade of the foreground color is mixed in.

```typescript
readonly gradientDegree: number;
```

Property Value:
- number

Remarks:
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### gradientStyle

Returns the gradient style for the fill.

```typescript
readonly gradientStyle: Word.GradientStyle | "Mixed" | "Horizontal" | "Vertical" | "DiagonalUp" | "DiagonalDown" | "FromCorner" | "FromTitle" | "FromCenter";
```

Property Value:
- [Word.GradientStyle](/en-us/javascript/api/word/word.gradientstyle) | "Mixed" | "Horizontal" | "Vertical" | "DiagonalUp" | "DiagonalDown" | "FromCorner" | "FromTitle" | "FromCenter"

Remarks:
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### gradientVariant

Returns the gradient variant for the fill as an integer value from 1 to 4 for most gradient fills.

```typescript
readonly gradientVariant: number;
```

Property Value:
- number

Remarks:
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### isVisible

Specifies if the object, or the formatting applied to it, is visible.

```typescript
isVisible: boolean;
```

Property Value:
- boolean

Remarks:
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### pattern

Returns a PatternType value that represents the pattern applied to the fill or line.

```typescript
readonly pattern: Word.PatternType | "Mixed" | "Percent5" | "Percent10" | "Percent20" | "Percent25" | "Percent30" | "Percent40" | "Percent50" | "Percent60" | "Percent70" | "Percent75" | "Percent80" | "Percent90" | "DarkHorizontal" | "DarkVertical" | "DarkDownwardDiagonal" | "DarkUpwardDiagonal" | "SmallCheckerBoard" | "Trellis" | "LightHorizontal" | "LightVertical" | "LightDownwardDiagonal" | "LightUpwardDiagonal" | "SmallGrid" | "DottedDiamond" | "WideDownwardDiagonal" | "WideUpwardDiagonal" | "DashedUpwardDiagonal" | "DashedDownwardDiagonal" | "NarrowVertical" | "NarrowHorizontal" | "DashedVertical" | "DashedHorizontal" | "LargeConfetti" | "LargeGrid" | "HorizontalBrick" | "LargeCheckerBoard" | "SmallConfetti" | "ZigZag" | "SolidDiamond" | "DiagonalBrick" | "OutlinedDiamond" | "Plaid" | "Sphere" | "Weave" | "DottedGrid" | "Divot" | "Shingle" | "Wave" | "Horizontal" | "Vertical" | "Cross" | "DownwardDiagonal" | "UpwardDiagonal" | "DiagonalCross";
```

Property Value:
- [Word.PatternType](/en-us/javascript/api/word/word.patterntype) | "Mixed" | "Percent5" | "Percent10" | "Percent20" | "Percent25" | "Percent30" | "Percent40" | "Percent50" | "Percent60" | "Percent70" | "Percent75" | "Percent80" | "Percent90" | "DarkHorizontal" | "DarkVertical" | "DarkDownwardDiagonal" | "DarkUpwardDiagonal" | "SmallCheckerBoard" | "Trellis" | "LightHorizontal" | "LightVertical" | "LightDownwardDiagonal" | "LightUpwardDiagonal" | "SmallGrid" | "DottedDiamond" | "WideDownwardDiagonal" | "WideUpwardDiagonal" | "DashedUpwardDiagonal" | "DashedDownwardDiagonal" | "NarrowVertical" | "NarrowHorizontal" | "DashedVertical" | "DashedHorizontal" | "LargeConfetti" | "LargeGrid" | "HorizontalBrick" | "LargeCheckerBoard" | "SmallConfetti" | "ZigZag" | "SolidDiamond" | "DiagonalBrick" | "OutlinedDiamond" | "Plaid" | "Sphere" | "Weave" | "DottedGrid" | "Divot" | "Shingle" | "Wave" | "Horizontal" | "Vertical" | "Cross" | "DownwardDiagonal" | "UpwardDiagonal" | "DiagonalCross"

Remarks:
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### presetGradientType

Returns the preset gradient type for the fill.

```typescript
readonly presetGradientType: Word.PresetGradientType | "Mixed" | "EarlySunset" | "LateSunset" | "Nightfall" | "Daybreak" | "Horizon" | "Desert" | "Ocean" | "CalmWater" | "Fire" | "Fog" | "Moss" | "Peacock" | "Wheat" | "Parchment" | "Mahogany" | "Rainbow" | "RainbowII" | "Gold" | "GoldII" | "Brass" | "Chrome" | "ChromeII" | "Silver" | "Sapphire";
```

Property Value:
- [Word.PresetGradientType](/en-us/javascript/api/word/word.presetgradienttype) | "Mixed" | "EarlySunset" | "LateSunset" | "Nightfall" | "Daybreak" | "Horizon" | "Desert" | "Ocean" | "CalmWater" | "Fire" | "Fog" | "Moss" | "Peacock" | "Wheat" | "Parchment" | "Mahogany" | "Rainbow" | "RainbowII" | "Gold" | "GoldII" | "Brass" | "Chrome" | "ChromeII" | "Silver" | "Sapphire"

Remarks:
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### presetTexture

Gets the preset texture.

```typescript
readonly presetTexture: Word.PresetTexture | "Mixed" | "Papyrus" | "Canvas" | "Denim" | "WovenMat" | "WaterDroplets" | "PaperBag" | "FishFossil" | "Sand" | "GreenMarble" | "WhiteMarble" | "BrownMarble" | "Granite" | "Newsprint" | "RecycledPaper" | "Parchment" | "Stationery" | "BlueTissuePaper" | "PinkTissuePaper" | "PurpleMesh" | "Bouquet" | "Cork" | "Walnut" | "Oak" | "MediumWood";
```

Property Value:
- [Word.PresetTexture](/en-us/javascript/api/word/word.presettexture) | "Mixed" | "Papyrus" | "Canvas" | "Denim" | "WovenMat" | "WaterDroplets" | "PaperBag" | "FishFossil" | "Sand" | "GreenMarble" | "WhiteMarble" | "BrownMarble" | "Granite" | "Newsprint" | "RecycledPaper" | "Parchment" | "Stationery" | "BlueTissuePaper" | "PinkTissuePaper" | "PurpleMesh" | "Bouquet" | "Cork" | "Walnut" | "Oak" | "MediumWood"

Remarks:
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### rotateWithObject

Specifies whether the fill rotates with the shape.

```typescript
rotateWithObject: boolean;
```

Property Value:
- boolean

Remarks:
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### textureAlignment

Specifies the alignment (the origin of the coordinate grid) for the tiling of the texture fill.

```typescript
textureAlignment: Word.TextureAlignment | "Mixed" | "TopLeft" | "Top" | "TopRight" | "Left" | "Center" | "Right" | "BottomLeft" | "Bottom" | "BottomRight";
```

Property Value:
- [Word.TextureAlignment](/en-us/javascript/api/word/word.texturealignment) | "Mixed" | "TopLeft" | "Top" | "TopRight" | "Left" | "Center" | "Right" | "BottomLeft" | "Bottom" | "BottomRight"

Remarks:
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### textureHorizontalScale

Specifies the horizontal scaling factor for the texture fill.

```typescript
textureHorizontalScale: number;
```

Property Value:
- number

Remarks:
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### textureName

Returns the name of the custom texture file for the fill.

```typescript
readonly textureName: string;
```

Property Value:
- string

Remarks:
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### textureOffsetX

Specifies the horizontal offset of the texture from the origin in points.

```typescript
textureOffsetX: number;
```

Property Value:
- number

Remarks:
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### textureOffsetY

Specifies the vertical offset of the texture.

```typescript
textureOffsetY: number;
```

Property Value:
- number

Remarks:
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### textureTile

Specifies whether the texture is tiled.

```typescript
textureTile: boolean;
```

Property Value:
- boolean

Remarks:
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### textureType

Returns the texture type for the fill.

```typescript
readonly textureType: Word.TextureType | "Mixed" | "Preset" | "UserDefined";
```

Property Value:
- [Word.TextureType](/en-us/javascript/api/word/word.texturetype) | "Mixed" | "Preset" | "UserDefined"

Remarks:
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### textureVerticalScale

Specifies the vertical scaling factor for the texture fill as a value between 0.0 and 1.0.

```typescript
textureVerticalScale: number;
```

Property Value:
- number

Remarks:
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### transparency

Specifies the degree of transparency of the fill for a shape as a value between 0.0 (opaque) and 1.0 (clear).

```typescript
transparency: number;
```

Property Value:
- number

Remarks:
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### type

Gets the fill format type.

```typescript
readonly type: Word.FillType | "Mixed" | "Solid" | "Patterned" | "Gradient" | "Textured" | "Background" | "Picture";
```

Property Value:
- [Word.FillType](/en-us/javascript/api/word/word.filltype) | "Mixed" | "Solid" | "Patterned" | "Gradient" | "Textured" | "Background" | "Picture"

Remarks:
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Method Details

### load(options)

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(options?: Word.Interfaces.FillFormatLoadOptions): Word.FillFormat;
```

Parameters:
- options: [Word.Interfaces.FillFormatLoadOptions](/en-us/javascript/api/word/word.interfaces.fillformatloadoptions)  
  Provides options for which properties of the object to load.

Returns:
- [Word.FillFormat](/en-us/javascript/api/word/word.fillformat)

### load(propertyNames)

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.FillFormat;
```

Parameters:
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns:
- [Word.FillFormat](/en-us/javascript/api/word/word.fillformat)

### load(propertyNamesAndPaths)

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
    select?: string;
    expand?: string;
}): Word.FillFormat;
```

Parameters:
- propertyNamesAndPaths: { select?: string; expand?: string; }  
  propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

Returns:
- [Word.FillFormat](/en-us/javascript/api/word/word.fillformat)

### set(properties, options)

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.FillFormatUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

Parameters:
- properties: [Word.Interfaces.FillFormatUpdateData](/en-us/javascript/api/word/word.interfaces.fillformatupdatedata)  
  A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
- options: [OfficeExtension.UpdateOptions](/en-us/javascript/api/office/officeextension.updateoptions)  
  Provides an option to suppress errors if the properties object tries to set any read-only properties.

Returns:
- void

### set(properties)

Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.FillFormat): void;
```

Parameters:
- properties: [Word.FillFormat](/en-us/javascript/api/word/word.fillformat)

Returns:
- void

### setOneColorGradient(style, variant, degree) {#setonecolorgradientstyle-variant-degree-1}

Sets the fill to a one-color gradient.

```typescript
setOneColorGradient(style: Word.GradientStyle, variant: number, degree: number): void;
```

Parameters:
- style: [Word.GradientStyle](/en-us/javascript/api/word/word.gradientstyle)  
  The gradient style.
- variant: number  
  The gradient variant. Can be a value from 1 to 4.
- degree: number  
  The gradient degree. Can be a value from 0.0 (dark) to 1.0 (light).

Returns:
- void

Remarks:
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### setOneColorGradient(style, variant, degree) {#setonecolorgradientstyle-variant-degree-2}

Sets the fill to a one-color gradient.

```typescript
setOneColorGradient(style: "Mixed" | "Horizontal" | "Vertical" | "DiagonalUp" | "DiagonalDown" | "FromCorner" | "FromTitle" | "FromCenter", variant: number, degree: number): void;
```

Parameters:
- style: "Mixed" | "Horizontal" | "Vertical" | "DiagonalUp" | "DiagonalDown" | "FromCorner" | "FromTitle" | "FromCenter"  
  The gradient style.
- variant: number  
  The gradient variant. Can be a value from 1 to 4.
- degree: number  
  The gradient degree. Can be a value from 0.0 (dark) to 1.0 (light).

Returns:
- void

Remarks:
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### setPatterned(pattern) {#setpatternedpattern-1}

Sets the fill to a pattern.

```typescript
setPatterned(pattern: Word.PatternType): void;
```

Parameters:
- pattern: [Word.PatternType](/en-us/javascript/api/word/word.patterntype)

Returns:
- void

Remarks:
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### setPatterned(pattern) {#setpatternedpattern-2}

Sets the fill to a pattern.

```typescript
setPatterned(pattern: "Mixed" | "Percent5" | "Percent10" | "Percent20" | "Percent25" | "Percent30" | "Percent40" | "Percent50" | "Percent60" | "Percent70" | "Percent75" | "Percent80" | "Percent90" | "DarkHorizontal" | "DarkVertical" | "DarkDownwardDiagonal" | "DarkUpwardDiagonal" | "SmallCheckerBoard" | "Trellis" | "LightHorizontal" | "LightVertical" | "LightDownwardDiagonal" | "LightUpwardDiagonal" | "SmallGrid" | "DottedDiamond" | "WideDownwardDiagonal" | "WideUpwardDiagonal" | "DashedUpwardDiagonal" | "DashedDownwardDiagonal" | "NarrowVertical" | "NarrowHorizontal" | "DashedVertical" | "DashedHorizontal" | "LargeConfetti" | "LargeGrid" | "HorizontalBrick" | "LargeCheckerBoard" | "SmallConfetti" | "ZigZag" | "SolidDiamond" | "DiagonalBrick" | "OutlinedDiamond" | "Plaid" | "Sphere" | "Weave" | "DottedGrid" | "Divot" | "Shingle" | "Wave" | "Horizontal" | "Vertical" | "Cross" | "DownwardDiagonal" | "UpwardDiagonal" | "DiagonalCross"): void;
```

Parameters:
- pattern: "Mixed" | "Percent5" | "Percent10" | "Percent20" | "Percent25" | "Percent30" | "Percent40" | "Percent50" | "Percent60" | "Percent70" | "Percent75" | "Percent80" | "Percent90" | "DarkHorizontal" | "DarkVertical" | "DarkDownwardDiagonal" | "DarkUpwardDiagonal" | "SmallCheckerBoard" | "Trellis" | "LightHorizontal" | "LightVertical" | "LightDownwardDiagonal" | "LightUpwardDiagonal" | "SmallGrid" | "DottedDiamond" | "WideDownwardDiagonal" | "WideUpwardDiagonal" | "DashedUpwardDiagonal" | "DashedDownwardDiagonal" | "NarrowVertical" | "NarrowHorizontal" | "DashedVertical" | "DashedHorizontal" | "LargeConfetti" | "LargeGrid" | "HorizontalBrick" | "LargeCheckerBoard" | "SmallConfetti" | "ZigZag" | "SolidDiamond" | "DiagonalBrick" | "OutlinedDiamond" | "Plaid" | "Sphere" | "Weave" | "DottedGrid" | "Divot" | "Shingle" | "Wave" | "Horizontal" | "Vertical" | "Cross" | "DownwardDiagonal" | "UpwardDiagonal" | "DiagonalCross"

Returns:
- void

Remarks:
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### setPresetGradient(style, variant, presetGradientType) {#setpresetgradientstyle-variant-presetgradienttype-1}

Sets the fill to a preset gradient. The gradient style. The gradient variant. Can be a value from 1 to 4. The preset gradient type.

```typescript
setPresetGradient(style: Word.GradientStyle, variant: number, presetGradientType: Word.PresetGradientType): void;
```

Parameters:
- style: [Word.GradientStyle](/en-us/javascript/api/word/word.gradientstyle)
- variant: number
- presetGradientType: [Word.PresetGradientType](/en-us/javascript/api/word/word.presetgradienttype)

Returns:
- void

Remarks:
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### setPresetGradient(style, variant, presetGradientType) {#setpresetgradientstyle-variant-presetgradienttype-2}

Sets the fill to a preset gradient. The gradient style. The gradient variant. Can be a value from 1 to 4. The preset gradient type.

```typescript
setPresetGradient(style: "Mixed" | "Horizontal" | "Vertical" | "DiagonalUp" | "DiagonalDown" | "FromCorner" | "FromTitle" | "FromCenter", variant: number, presetGradientType: "Mixed" | "EarlySunset" | "LateSunset" | "Nightfall" | "Daybreak" | "Horizon" | "Desert" | "Ocean" | "CalmWater" | "Fire" | "Fog" | "Moss" | "Peacock" | "Wheat" | "Parchment" | "Mahogany" | "Rainbow" | "RainbowII" | "Gold" | "GoldII" | "Brass" | "Chrome" | "ChromeII" | "Silver" | "Sapphire"): void;
```

Parameters:
- style: "Mixed" | "Horizontal" | "Vertical" | "DiagonalUp" | "DiagonalDown" | "FromCorner" | "FromTitle" | "FromCenter"
- variant: number
- presetGradientType: "Mixed" | "EarlySunset" | "LateSunset" | "Nightfall" | "Daybreak" | "Horizon" | "Desert" | "Ocean" | "CalmWater" | "Fire" | "Fog" | "Moss" | "Peacock" | "Wheat" | "Parchment" | "Mahogany" | "Rainbow" | "RainbowII" | "Gold" | "GoldII" | "Brass" | "Chrome" | "ChromeII" | "Silver" | "Sapphire"

Returns:
- void

Remarks:
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### setPresetTextured(presetTexture) {#setpresettexturedpresettexture-1}

Sets the fill to a preset texture.

```typescript
setPresetTextured(presetTexture: Word.PresetTexture): void;
```

Parameters:
- presetTexture: [Word.PresetTexture](/en-us/javascript/api/word/word.presettexture)

Returns:
- void

Remarks:
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### setPresetTextured(presetTexture) {#setpresettexturedpresettexture-2}

Sets the fill to a preset texture.

```typescript
setPresetTextured(presetTexture: "Mixed" | "Papyrus" | "Canvas" | "Denim" | "WovenMat" | "WaterDroplets" | "PaperBag" | "FishFossil" | "Sand" | "GreenMarble" | "WhiteMarble" | "BrownMarble" | "Granite" | "Newsprint" | "RecycledPaper" | "Parchment" | "Stationery" | "BlueTissuePaper" | "PinkTissuePaper" | "PurpleMesh" | "Bouquet" | "Cork" | "Walnut" | "Oak" | "MediumWood"): void;
```

Parameters:
- presetTexture: "Mixed" | "Papyrus" | "Canvas" | "Denim" | "WovenMat" | "WaterDroplets" | "PaperBag" | "FishFossil" | "Sand" | "GreenMarble" | "WhiteMarble" | "BrownMarble" | "Granite" | "Newsprint" | "RecycledPaper" | "Parchment" | "Stationery" | "BlueTissuePaper" | "PinkTissuePaper" | "PurpleMesh" | "Bouquet" | "Cork" | "Walnut" | "Oak" | "MediumWood"

Returns:
- void

Remarks:
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### setTwoColorGradient(style, variant) {#settwocolorgradientstyle-variant-1}

Sets the fill to a two-color gradient.

```typescript
setTwoColorGradient(style: Word.GradientStyle, variant: number): void;
```

Parameters:
- style: [Word.GradientStyle](/en-us/javascript/api/word/word.gradientstyle)
- variant: number

Returns:
- void

Remarks:
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### setTwoColorGradient(style, variant) {#settwocolorgradientstyle-variant-2}

Sets the fill to a two-color gradient.

```typescript
setTwoColorGradient(style: "Mixed" | "Horizontal" | "Vertical" | "DiagonalUp" | "DiagonalDown" | "FromCorner" | "FromTitle" | "FromCenter", variant: number): void;
```

Parameters:
- style: "Mixed" | "Horizontal" | "Vertical" | "DiagonalUp" | "DiagonalDown" | "FromCorner" | "FromTitle" | "FromCenter"
- variant: number

Returns:
- void

Remarks:
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### solid

Sets the fill to a uniform color.

```typescript
solid(): void;
```

Returns:
- void

Remarks:
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### toJSON

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.FillFormat object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.FillFormatData) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.FillFormatData;
```

Returns:
- [Word.Interfaces.FillFormatData](/en-us/javascript/api/word/word.interfaces.fillformatdata)

### track

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.FillFormat;
```

Returns:
- [Word.FillFormat](/en-us/javascript/api/word/word.fillformat)

### untrack

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.FillFormat;
```

Returns:
- [Word.FillFormat](/en-us/javascript/api/word/word.fillformat)