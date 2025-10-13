# Word.Interfaces.ThreeDimensionalFormatUpdateData interface

Package: [word](/en-us/javascript/api/word)

An interface for updating data on the ThreeDimensionalFormat object, for use in threeDimensionalFormat.set({ ... }).

## Properties

- [bevelBottomDepth](#bevelbottomdepth) - Specifies the depth of the bottom bevel.
- [bevelBottomInset](#bevelbottominset) - Specifies the inset size for the bottom bevel.
- [bevelBottomType](#bevelbottomtype) - Specifies a BevelType value that represents the bevel type for the bottom bevel.
- [bevelTopDepth](#beveltopdepth) - Specifies the depth of the top bevel.
- [bevelTopInset](#beveltopinset) - Specifies the inset size for the top bevel.
- [bevelTopType](#beveltoptype) - Specifies a BevelType value that represents the bevel type for the top bevel.
- [contourColor](#contourcolor) - Returns a ColorFormat object that represents color of the contour of a shape.
- [contourWidth](#contourwidth) - Specifies the width of the contour of a shape.
- [depth](#depth) - Specifies the depth of the shape's extrusion.
- [extrusionColor](#extrusioncolor) - Returns a ColorFormat object that represents the color of the shape's extrusion.
- [extrusionColorType](#extrusioncolortype) - Specifies whether the extrusion color is based on the extruded shape's fill (the front face of the extrusion) and automatically changes when the shape's fill changes, or whether the extrusion color is independent of the shape's fill.
- [fieldOfView](#fieldofview) - Specifies the amount of perspective for a shape.
- [isPerspective](#isperspective) - Specifies true if the extrusion appears in perspective â that is, if the walls of the extrusion narrow toward a vanishing point, false if the extrusion is a parallel, or orthographic, projection â that is, if the walls don't narrow toward a vanishing point.
- [isVisible](#isvisible) - Specifies if the specified object, or the formatting applied to it, is visible.
- [lightAngle](#lightangle) - Specifies the angle of the lighting.
- [presetLighting](#presetlighting) - Specifies a LightRigType value that represents the lighting preset.
- [presetLightingDirection](#presetlightingdirection) - Specifies the position of the light source relative to the extrusion.
- [presetLightingSoftness](#presetlightingsoftness) - Specifies the intensity of the extrusion lighting.
- [presetMaterial](#presetmaterial) - Specifies the extrusion surface material.
- [projectText](#projecttext) - Specifies whether text on a shape rotates with shape. true rotates the text.
- [rotationX](#rotationx) - Specifies the rotation of the extruded shape around the x-axis in degrees.
- [rotationY](#rotationy) - Specifies the rotation of the extruded shape around the y-axis in degrees.
- [rotationZ](#rotationz) - Specifies the z-axis rotation of the camera.
- [z](#z) - Specifies the position on the z-axis for the shape.

## Property Details

### bevelBottomDepth

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the depth of the bottom bevel.

```typescript
bevelBottomDepth?: number;
```

Property Value
- number

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### bevelBottomInset

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the inset size for the bottom bevel.

```typescript
bevelBottomInset?: number;
```

Property Value
- number

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### bevelBottomType

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a BevelType value that represents the bevel type for the bottom bevel.

```typescript
bevelBottomType?: Word.BevelType | "mixed" | "none" | "relaxedInset" | "circle" | "slope" | "cross" | "angle" | "softRound" | "convex" | "coolSlant" | "divot" | "riblet" | "hardEdge" | "artDeco";
```

Property Value
- [Word.BevelType](/en-us/javascript/api/word/word.beveltype) | "mixed" | "none" | "relaxedInset" | "circle" | "slope" | "cross" | "angle" | "softRound" | "convex" | "coolSlant" | "divot" | "riblet" | "hardEdge" | "artDeco"

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### bevelTopDepth

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the depth of the top bevel.

```typescript
bevelTopDepth?: number;
```

Property Value
- number

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### bevelTopInset

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the inset size for the top bevel.

```typescript
bevelTopInset?: number;
```

Property Value
- number

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### bevelTopType

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a BevelType value that represents the bevel type for the top bevel.

```typescript
bevelTopType?: Word.BevelType | "mixed" | "none" | "relaxedInset" | "circle" | "slope" | "cross" | "angle" | "softRound" | "convex" | "coolSlant" | "divot" | "riblet" | "hardEdge" | "artDeco";
```

Property Value
- [Word.BevelType](/en-us/javascript/api/word/word.beveltype) | "mixed" | "none" | "relaxedInset" | "circle" | "slope" | "cross" | "angle" | "softRound" | "convex" | "coolSlant" | "divot" | "riblet" | "hardEdge" | "artDeco"

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### contourColor

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a ColorFormat object that represents color of the contour of a shape.

```typescript
contourColor?: Word.Interfaces.ColorFormatUpdateData;
```

Property Value
- [Word.Interfaces.ColorFormatUpdateData](/en-us/javascript/api/word/word.interfaces.colorformatupdatedata)

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### contourWidth

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the width of the contour of a shape.

```typescript
contourWidth?: number;
```

Property Value
- number

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### depth

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the depth of the shape's extrusion.

```typescript
depth?: number;
```

Property Value
- number

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### extrusionColor

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a ColorFormat object that represents the color of the shape's extrusion.

```typescript
extrusionColor?: Word.Interfaces.ColorFormatUpdateData;
```

Property Value
- [Word.Interfaces.ColorFormatUpdateData](/en-us/javascript/api/word/word.interfaces.colorformatupdatedata)

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### extrusionColorType

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the extrusion color is based on the extruded shape's fill (the front face of the extrusion) and automatically changes when the shape's fill changes, or whether the extrusion color is independent of the shape's fill.

```typescript
extrusionColorType?: Word.ExtrusionColorType | "mixed" | "automatic" | "custom";
```

Property Value
- [Word.ExtrusionColorType](/en-us/javascript/api/word/word.extrusioncolortype) | "mixed" | "automatic" | "custom"

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### fieldOfView

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the amount of perspective for a shape.

```typescript
fieldOfView?: number;
```

Property Value
- number

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### isPerspective

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies true if the extrusion appears in perspective â that is, if the walls of the extrusion narrow toward a vanishing point, false if the extrusion is a parallel, or orthographic, projection â that is, if the walls don't narrow toward a vanishing point.

```typescript
isPerspective?: boolean;
```

Property Value
- boolean

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### isVisible

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the specified object, or the formatting applied to it, is visible.

```typescript
isVisible?: boolean;
```

Property Value
- boolean

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### lightAngle

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the angle of the lighting.

```typescript
lightAngle?: number;
```

Property Value
- number

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### presetLighting

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a LightRigType value that represents the lighting preset.

```typescript
presetLighting?: Word.LightRigType | "Mixed" | "LegacyFlat1" | "LegacyFlat2" | "LegacyFlat3" | "LegacyFlat4" | "LegacyNormal1" | "LegacyNormal2" | "LegacyNormal3" | "LegacyNormal4" | "LegacyHarsh1" | "LegacyHarsh2" | "LegacyHarsh3" | "LegacyHarsh4" | "ThreePoint" | "Balanced" | "Soft" | "Harsh" | "Flood" | "Contrasting" | "Morning" | "Sunrise" | "Sunset" | "Chilly" | "Freezing" | "Flat" | "TwoPoint" | "Glow" | "BrightRoom";
```

Property Value
- [Word.LightRigType](/en-us/javascript/api/word/word.lightrigtype) | "Mixed" | "LegacyFlat1" | "LegacyFlat2" | "LegacyFlat3" | "LegacyFlat4" | "LegacyNormal1" | "LegacyNormal2" | "LegacyNormal3" | "LegacyNormal4" | "LegacyHarsh1" | "LegacyHarsh2" | "LegacyHarsh3" | "LegacyHarsh4" | "ThreePoint" | "Balanced" | "Soft" | "Harsh" | "Flood" | "Contrasting" | "Morning" | "Sunrise" | "Sunset" | "Chilly" | "Freezing" | "Flat" | "TwoPoint" | "Glow" | "BrightRoom"

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### presetLightingDirection

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the position of the light source relative to the extrusion.

```typescript
presetLightingDirection?: Word.PresetLightingDirection | "Mixed" | "TopLeft" | "Top" | "TopRight" | "Left" | "None" | "Right" | "BottomLeft" | "Bottom" | "BottomRight";
```

Property Value
- [Word.PresetLightingDirection](/en-us/javascript/api/word/word.presetlightingdirection) | "Mixed" | "TopLeft" | "Top" | "TopRight" | "Left" | "None" | "Right" | "BottomLeft" | "Bottom" | "BottomRight"

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### presetLightingSoftness

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the intensity of the extrusion lighting.

```typescript
presetLightingSoftness?: Word.PresetLightingSoftness | "Mixed" | "Dim" | "Normal" | "Bright";
```

Property Value
- [Word.PresetLightingSoftness](/en-us/javascript/api/word/word.presetlightingsoftness) | "Mixed" | "Dim" | "Normal" | "Bright"

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### presetMaterial

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the extrusion surface material.

```typescript
presetMaterial?: Word.PresetMaterial | "Mixed" | "Matte" | "Plastic" | "Metal" | "WireFrame" | "Matte2" | "Plastic2" | "Metal2" | "WarmMatte" | "TranslucentPowder" | "Powder" | "DarkEdge" | "SoftEdge" | "Clear" | "Flat" | "SoftMetal";
```

Property Value
- [Word.PresetMaterial](/en-us/javascript/api/word/word.presetmaterial) | "Mixed" | "Matte" | "Plastic" | "Metal" | "WireFrame" | "Matte2" | "Plastic2" | "Metal2" | "WarmMatte" | "TranslucentPowder" | "Powder" | "DarkEdge" | "SoftEdge" | "Clear" | "Flat" | "SoftMetal"

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### projectText

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether text on a shape rotates with shape. true rotates the text.

```typescript
projectText?: boolean;
```

Property Value
- boolean

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### rotationX

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the rotation of the extruded shape around the x-axis in degrees.

```typescript
rotationX?: number;
```

Property Value
- number

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### rotationY

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the rotation of the extruded shape around the y-axis in degrees.

```typescript
rotationY?: number;
```

Property Value
- number

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### rotationZ

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the z-axis rotation of the camera.

```typescript
rotationZ?: number;
```

Property Value
- number

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### z

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the position on the z-axis for the shape.

```typescript
z?: number;
```

Property Value
- number

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)