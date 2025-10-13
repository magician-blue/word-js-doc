# Word.Interfaces.ThreeDimensionalFormatData interface

Package: [word](/en-us/javascript/api/word)

An interface describing the data returned by calling `threeDimensionalFormat.toJSON()`.

## Properties

| Property | Description |
| --- | --- |
| bevelBottomDepth | Specifies the depth of the bottom bevel. |
| bevelBottomInset | Specifies the inset size for the bottom bevel. |
| bevelBottomType | Specifies a `BevelType` value that represents the bevel type for the bottom bevel. |
| bevelTopDepth | Specifies the depth of the top bevel. |
| bevelTopInset | Specifies the inset size for the top bevel. |
| bevelTopType | Specifies a `BevelType` value that represents the bevel type for the top bevel. |
| contourColor | Returns a `ColorFormat` object that represents color of the contour of a shape. |
| contourWidth | Specifies the width of the contour of a shape. |
| depth | Specifies the depth of the shape's extrusion. |
| extrusionColor | Returns a `ColorFormat` object that represents the color of the shape's extrusion. |
| extrusionColorType | Specifies whether the extrusion color is based on the extruded shape's fill (the front face of the extrusion) and automatically changes when the shape's fill changes, or whether the extrusion color is independent of the shape's fill. |
| fieldOfView | Specifies the amount of perspective for a shape. |
| isPerspective | Specifies `true` if the extrusion appears in perspective â that is, if the walls of the extrusion narrow toward a vanishing point, `false` if the extrusion is a parallel, or orthographic, projection â that is, if the walls don't narrow toward a vanishing point. |
| isVisible | Specifies if the specified object, or the formatting applied to it, is visible. |
| lightAngle | Specifies the angle of the lighting. |
| presetCamera | Returns a `PresetCamera` value that represents the camera presets. |
| presetExtrusionDirection | Returns the direction taken by the extrusion's sweep path leading away from the extruded shape (the front face of the extrusion). |
| presetLighting | Specifies a `LightRigType` value that represents the lighting preset. |
| presetLightingDirection | Specifies the position of the light source relative to the extrusion. |
| presetLightingSoftness | Specifies the intensity of the extrusion lighting. |
| presetMaterial | Specifies the extrusion surface material. |
| presetThreeDimensionalFormat | Returns the preset extrusion format. |
| projectText | Specifies whether text on a shape rotates with shape. `true` rotates the text. |
| rotationX | Specifies the rotation of the extruded shape around the x-axis in degrees. |
| rotationY | Specifies the rotation of the extruded shape around the y-axis in degrees. |
| rotationZ | Specifies the z-axis rotation of the camera. |
| z | Specifies the position on the z-axis for the shape. |

## Property Details

### bevelBottomDepth

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the depth of the bottom bevel.

```typescript
bevelBottomDepth?: number;
```

#### Property Value
number

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### bevelBottomInset

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the inset size for the bottom bevel.

```typescript
bevelBottomInset?: number;
```

#### Property Value
number

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### bevelBottomType

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a `BevelType` value that represents the bevel type for the bottom bevel.

```typescript
bevelBottomType?: Word.BevelType | "mixed" | "none" | "relaxedInset" | "circle" | "slope" | "cross" | "angle" | "softRound" | "convex" | "coolSlant" | "divot" | "riblet" | "hardEdge" | "artDeco";
```

#### Property Value
[Word.BevelType](/en-us/javascript/api/word/word.beveltype) | "mixed" | "none" | "relaxedInset" | "circle" | "slope" | "cross" | "angle" | "softRound" | "convex" | "coolSlant" | "divot" | "riblet" | "hardEdge" | "artDeco"

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### bevelTopDepth

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the depth of the top bevel.

```typescript
bevelTopDepth?: number;
```

#### Property Value
number

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### bevelTopInset

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the inset size for the top bevel.

```typescript
bevelTopInset?: number;
```

#### Property Value
number

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### bevelTopType

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a `BevelType` value that represents the bevel type for the top bevel.

```typescript
bevelTopType?: Word.BevelType | "mixed" | "none" | "relaxedInset" | "circle" | "slope" | "cross" | "angle" | "softRound" | "convex" | "coolSlant" | "divot" | "riblet" | "hardEdge" | "artDeco";
```

#### Property Value
[Word.BevelType](/en-us/javascript/api/word/word.beveltype) | "mixed" | "none" | "relaxedInset" | "circle" | "slope" | "cross" | "angle" | "softRound" | "convex" | "coolSlant" | "divot" | "riblet" | "hardEdge" | "artDeco"

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### contourColor

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `ColorFormat` object that represents color of the contour of a shape.

```typescript
contourColor?: Word.Interfaces.ColorFormatData;
```

#### Property Value
[Word.Interfaces.ColorFormatData](/en-us/javascript/api/word/word.interfaces.colorformatdata)

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### contourWidth

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the width of the contour of a shape.

```typescript
contourWidth?: number;
```

#### Property Value
number

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### depth

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the depth of the shape's extrusion.

```typescript
depth?: number;
```

#### Property Value
number

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### extrusionColor

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `ColorFormat` object that represents the color of the shape's extrusion.

```typescript
extrusionColor?: Word.Interfaces.ColorFormatData;
```

#### Property Value
[Word.Interfaces.ColorFormatData](/en-us/javascript/api/word/word.interfaces.colorformatdata)

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### extrusionColorType

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the extrusion color is based on the extruded shape's fill (the front face of the extrusion) and automatically changes when the shape's fill changes, or whether the extrusion color is independent of the shape's fill.

```typescript
extrusionColorType?: Word.ExtrusionColorType | "mixed" | "automatic" | "custom";
```

#### Property Value
[Word.ExtrusionColorType](/en-us/javascript/api/word/word.extrusioncolortype) | "mixed" | "automatic" | "custom"

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### fieldOfView

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the amount of perspective for a shape.

```typescript
fieldOfView?: number;
```

#### Property Value
number

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### isPerspective

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies `true` if the extrusion appears in perspective â that is, if the walls of the extrusion narrow toward a vanishing point, `false` if the extrusion is a parallel, or orthographic, projection â that is, if the walls don't narrow toward a vanishing point.

```typescript
isPerspective?: boolean;
```

#### Property Value
boolean

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### isVisible

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the specified object, or the formatting applied to it, is visible.

```typescript
isVisible?: boolean;
```

#### Property Value
boolean

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### lightAngle

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the angle of the lighting.

```typescript
lightAngle?: number;
```

#### Property Value
number

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### presetCamera

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `PresetCamera` value that represents the camera presets.

```typescript
presetCamera?: Word.PresetCamera | "Mixed" | "LegacyObliqueTopLeft" | "LegacyObliqueTop" | "LegacyObliqueTopRight" | "LegacyObliqueLeft" | "LegacyObliqueFront" | "LegacyObliqueRight" | "LegacyObliqueBottomLeft" | "LegacyObliqueBottom" | "LegacyObliqueBottomRight" | "LegacyPerspectiveTopLeft" | "LegacyPerspectiveTop" | "LegacyPerspectiveTopRight" | "LegacyPerspectiveLeft" | "LegacyPerspectiveFront" | "LegacyPerspectiveRight" | "LegacyPerspectiveBottomLeft" | "LegacyPerspectiveBottom" | "LegacyPerspectiveBottomRight" | "OrthographicFront" | "IsometricTopUp" | "IsometricTopDown" | "IsometricBottomUp" | "IsometricBottomDown" | "IsometricLeftUp" | "IsometricLeftDown" | "IsometricRightUp" | "IsometricRightDown" | "IsometricOffAxis1Left" | "IsometricOffAxis1Right" | "IsometricOffAxis1Top" | "IsometricOffAxis2Left" | "IsometricOffAxis2Right" | "IsometricOffAxis2Top" | "IsometricOffAxis3Left" | "IsometricOffAxis3Right" | "IsometricOffAxis3Bottom" | "IsometricOffAxis4Left" | "IsometricOffAxis4Right" | "IsometricOffAxis4Bottom" | "ObliqueTopLeft" | "ObliqueTop" | "ObliqueTopRight" | "ObliqueLeft" | "ObliqueRight" | "ObliqueBottomLeft" | "ObliqueBottom" | "ObliqueBottomRight" | "PerspectiveFront" | "PerspectiveLeft" | "PerspectiveRight" | "PerspectiveAbove" | "PerspectiveBelow" | "PerspectiveAboveLeftFacing" | "PerspectiveAboveRightFacing" | "PerspectiveContrastingLeftFacing" | "PerspectiveContrastingRightFacing" | "PerspectiveHeroicLeftFacing" | "PerspectiveHeroicRightFacing" | "PerspectiveHeroicExtremeLeftFacing" | "PerspectiveHeroicExtremeRightFacing" | "PerspectiveRelaxed" | "PerspectiveRelaxedModerately";
```

#### Property Value
[Word.PresetCamera](/en-us/javascript/api/word/word.presetcamera) | "Mixed" | "LegacyObliqueTopLeft" | "LegacyObliqueTop" | "LegacyObliqueTopRight" | "LegacyObliqueLeft" | "LegacyObliqueFront" | "LegacyObliqueRight" | "LegacyObliqueBottomLeft" | "LegacyObliqueBottom" | "LegacyObliqueBottomRight" | "LegacyPerspectiveTopLeft" | "LegacyPerspectiveTop" | "LegacyPerspectiveTopRight" | "LegacyPerspectiveLeft" | "LegacyPerspectiveFront" | "LegacyPerspectiveRight" | "LegacyPerspectiveBottomLeft" | "LegacyPerspectiveBottom" | "LegacyPerspectiveBottomRight" | "OrthographicFront" | "IsometricTopUp" | "IsometricTopDown" | "IsometricBottomUp" | "IsometricBottomDown" | "IsometricLeftUp" | "IsometricLeftDown" | "IsometricRightUp" | "IsometricRightDown" | "IsometricOffAxis1Left" | "IsometricOffAxis1Right" | "IsometricOffAxis1Top" | "IsometricOffAxis2Left" | "IsometricOffAxis2Right" | "IsometricOffAxis2Top" | "IsometricOffAxis3Left" | "IsometricOffAxis3Right" | "IsometricOffAxis3Bottom" | "IsometricOffAxis4Left" | "IsometricOffAxis4Right" | "IsometricOffAxis4Bottom" | "ObliqueTopLeft" | "ObliqueTop" | "ObliqueTopRight" | "ObliqueLeft" | "ObliqueRight" | "ObliqueBottomLeft" | "ObliqueBottom" | "ObliqueBottomRight" | "PerspectiveFront" | "PerspectiveLeft" | "PerspectiveRight" | "PerspectiveAbove" | "PerspectiveBelow" | "PerspectiveAboveLeftFacing" | "PerspectiveAboveRightFacing" | "PerspectiveContrastingLeftFacing" | "PerspectiveContrastingRightFacing" | "PerspectiveHeroicLeftFacing" | "PerspectiveHeroicRightFacing" | "PerspectiveHeroicExtremeLeftFacing" | "PerspectiveHeroicExtremeRightFacing" | "PerspectiveRelaxed" | "PerspectiveRelaxedModerately"

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### presetExtrusionDirection

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the direction taken by the extrusion's sweep path leading away from the extruded shape (the front face of the extrusion).

```typescript
presetExtrusionDirection?: Word.PresetExtrusionDirection | "Mixed" | "BottomRight" | "Bottom" | "BottomLeft" | "Right" | "None" | "Left" | "TopRight" | "Top" | "TopLeft";
```

#### Property Value
[Word.PresetExtrusionDirection](/en-us/javascript/api/word/word.presetextrusiondirection) | "Mixed" | "BottomRight" | "Bottom" | "BottomLeft" | "Right" | "None" | "Left" | "TopRight" | "Top" | "TopLeft"

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### presetLighting

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a `LightRigType` value that represents the lighting preset.

```typescript
presetLighting?: Word.LightRigType | "Mixed" | "LegacyFlat1" | "LegacyFlat2" | "LegacyFlat3" | "LegacyFlat4" | "LegacyNormal1" | "LegacyNormal2" | "LegacyNormal3" | "LegacyNormal4" | "LegacyHarsh1" | "LegacyHarsh2" | "LegacyHarsh3" | "LegacyHarsh4" | "ThreePoint" | "Balanced" | "Soft" | "Harsh" | "Flood" | "Contrasting" | "Morning" | "Sunrise" | "Sunset" | "Chilly" | "Freezing" | "Flat" | "TwoPoint" | "Glow" | "BrightRoom";
```

#### Property Value
[Word.LightRigType](/en-us/javascript/api/word/word.lightrigtype) | "Mixed" | "LegacyFlat1" | "LegacyFlat2" | "LegacyFlat3" | "LegacyFlat4" | "LegacyNormal1" | "LegacyNormal2" | "LegacyNormal3" | "LegacyNormal4" | "LegacyHarsh1" | "LegacyHarsh2" | "LegacyHarsh3" | "LegacyHarsh4" | "ThreePoint" | "Balanced" | "Soft" | "Harsh" | "Flood" | "Contrasting" | "Morning" | "Sunrise" | "Sunset" | "Chilly" | "Freezing" | "Flat" | "TwoPoint" | "Glow" | "BrightRoom"

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### presetLightingDirection

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the position of the light source relative to the extrusion.

```typescript
presetLightingDirection?: Word.PresetLightingDirection | "Mixed" | "TopLeft" | "Top" | "TopRight" | "Left" | "None" | "Right" | "BottomLeft" | "Bottom" | "BottomRight";
```

#### Property Value
[Word.PresetLightingDirection](/en-us/javascript/api/word/word.presetlightingdirection) | "Mixed" | "TopLeft" | "Top" | "TopRight" | "Left" | "None" | "Right" | "BottomLeft" | "Bottom" | "BottomRight"

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### presetLightingSoftness

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the intensity of the extrusion lighting.

```typescript
presetLightingSoftness?: Word.PresetLightingSoftness | "Mixed" | "Dim" | "Normal" | "Bright";
```

#### Property Value
[Word.PresetLightingSoftness](/en-us/javascript/api/word/word.presetlightingsoftness) | "Mixed" | "Dim" | "Normal" | "Bright"

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### presetMaterial

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the extrusion surface material.

```typescript
presetMaterial?: Word.PresetMaterial | "Mixed" | "Matte" | "Plastic" | "Metal" | "WireFrame" | "Matte2" | "Plastic2" | "Metal2" | "WarmMatte" | "TranslucentPowder" | "Powder" | "DarkEdge" | "SoftEdge" | "Clear" | "Flat" | "SoftMetal";
```

#### Property Value
[Word.PresetMaterial](/en-us/javascript/api/word/word.presetmaterial) | "Mixed" | "Matte" | "Plastic" | "Metal" | "WireFrame" | "Matte2" | "Plastic2" | "Metal2" | "WarmMatte" | "TranslucentPowder" | "Powder" | "DarkEdge" | "SoftEdge" | "Clear" | "Flat" | "SoftMetal"

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### presetThreeDimensionalFormat

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the preset extrusion format.

```typescript
presetThreeDimensionalFormat?: Word.PresetThreeDimensionalFormat | "Mixed" | "Format1" | "Format2" | "Format3" | "Format4" | "Format5" | "Format6" | "Format7" | "Format8" | "Format9" | "Format10" | "Format11" | "Format12" | "Format13" | "Format14" | "Format15" | "Format16" | "Format17" | "Format18" | "Format19" | "Format20";
```

#### Property Value
[Word.PresetThreeDimensionalFormat](/en-us/javascript/api/word/word.presetthreedimensionalformat) | "Mixed" | "Format1" | "Format2" | "Format3" | "Format4" | "Format5" | "Format6" | "Format7" | "Format8" | "Format9" | "Format10" | "Format11" | "Format12" | "Format13" | "Format14" | "Format15" | "Format16" | "Format17" | "Format18" | "Format19" | "Format20"

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### projectText

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether text on a shape rotates with shape. `true` rotates the text.

```typescript
projectText?: boolean;
```

#### Property Value
boolean

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### rotationX

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the rotation of the extruded shape around the x-axis in degrees.

```typescript
rotationX?: number;
```

#### Property Value
number

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### rotationY

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the rotation of the extruded shape around the y-axis in degrees.

```typescript
rotationY?: number;
```

#### Property Value
number

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### rotationZ

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the z-axis rotation of the camera.

```typescript
rotationZ?: number;
```

#### Property Value
number

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### z

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the position on the z-axis for the shape.

```typescript
z?: number;
```

#### Property Value
number

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)