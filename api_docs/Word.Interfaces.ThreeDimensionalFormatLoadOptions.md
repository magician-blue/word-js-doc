# Word.Interfaces.ThreeDimensionalFormatLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents a shape's three-dimensional formatting.

## Remarks

[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- [$all](#word-word-interfaces-threedimensionalformatloadoptions-all-member) — Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
- [bevelBottomDepth](#word-word-interfaces-threedimensionalformatloadoptions-bevelbottomdepth-member) — Specifies the depth of the bottom bevel.
- [bevelBottomInset](#word-word-interfaces-threedimensionalformatloadoptions-bevelbottominset-member) — Specifies the inset size for the bottom bevel.
- [bevelBottomType](#word-word-interfaces-threedimensionalformatloadoptions-bevelbottomtype-member) — Specifies a `BevelType` value that represents the bevel type for the bottom bevel.
- [bevelTopDepth](#word-word-interfaces-threedimensionalformatloadoptions-beveltopdepth-member) — Specifies the depth of the top bevel.
- [bevelTopInset](#word-word-interfaces-threedimensionalformatloadoptions-beveltopinset-member) — Specifies the inset size for the top bevel.
- [bevelTopType](#word-word-interfaces-threedimensionalformatloadoptions-beveltoptype-member) — Specifies a `BevelType` value that represents the bevel type for the top bevel.
- [contourColor](#word-word-interfaces-threedimensionalformatloadoptions-contourcolor-member) — Returns a `ColorFormat` object that represents color of the contour of a shape.
- [contourWidth](#word-word-interfaces-threedimensionalformatloadoptions-contourwidth-member) — Specifies the width of the contour of a shape.
- [depth](#word-word-interfaces-threedimensionalformatloadoptions-depth-member) — Specifies the depth of the shape's extrusion.
- [extrusionColor](#word-word-interfaces-threedimensionalformatloadoptions-extrusioncolor-member) — Returns a `ColorFormat` object that represents the color of the shape's extrusion.
- [extrusionColorType](#word-word-interfaces-threedimensionalformatloadoptions-extrusioncolortype-member) — Specifies whether the extrusion color is based on the extruded shape's fill (the front face of the extrusion) and automatically changes when the shape's fill changes, or whether the extrusion color is independent of the shape's fill.
- [fieldOfView](#word-word-interfaces-threedimensionalformatloadoptions-fieldofview-member) — Specifies the amount of perspective for a shape.
- [isPerspective](#word-word-interfaces-threedimensionalformatloadoptions-isperspective-member) — Specifies `true` if the extrusion appears in perspective â that is, if the walls of the extrusion narrow toward a vanishing point, `false` if the extrusion is a parallel, or orthographic, projection â that is, if the walls don't narrow toward a vanishing point.
- [isVisible](#word-word-interfaces-threedimensionalformatloadoptions-isvisible-member) — Specifies if the specified object, or the formatting applied to it, is visible.
- [lightAngle](#word-word-interfaces-threedimensionalformatloadoptions-lightangle-member) — Specifies the angle of the lighting.
- [presetCamera](#word-word-interfaces-threedimensionalformatloadoptions-presetcamera-member) — Returns a `PresetCamera` value that represents the camera presets.
- [presetExtrusionDirection](#word-word-interfaces-threedimensionalformatloadoptions-presetextrusiondirection-member) — Returns the direction taken by the extrusion's sweep path leading away from the extruded shape (the front face of the extrusion).
- [presetLighting](#word-word-interfaces-threedimensionalformatloadoptions-presetlighting-member) — Specifies a `LightRigType` value that represents the lighting preset.
- [presetLightingDirection](#word-word-interfaces-threedimensionalformatloadoptions-presetlightingdirection-member) — Specifies the position of the light source relative to the extrusion.
- [presetLightingSoftness](#word-word-interfaces-threedimensionalformatloadoptions-presetlightingsoftness-member) — Specifies the intensity of the extrusion lighting.
- [presetMaterial](#word-word-interfaces-threedimensionalformatloadoptions-presetmaterial-member) — Specifies the extrusion surface material.
- [presetThreeDimensionalFormat](#word-word-interfaces-threedimensionalformatloadoptions-presetthreedimensionalformat-member) — Returns the preset extrusion format.
- [projectText](#word-word-interfaces-threedimensionalformatloadoptions-projecttext-member) — Specifies whether text on a shape rotates with shape. `true` rotates the text.
- [rotationX](#word-word-interfaces-threedimensionalformatloadoptions-rotationx-member) — Specifies the rotation of the extruded shape around the x-axis in degrees.
- [rotationY](#word-word-interfaces-threedimensionalformatloadoptions-rotationy-member) — Specifies the rotation of the extruded shape around the y-axis in degrees.
- [rotationZ](#word-word-interfaces-threedimensionalformatloadoptions-rotationz-member) — Specifies the z-axis rotation of the camera.
- [z](#word-word-interfaces-threedimensionalformatloadoptions-z-member) — Specifies the position on the z-axis for the shape.

## Property Details

### $all
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

Property Value
boolean

---

### bevelBottomDepth
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the depth of the bottom bevel.

```typescript
bevelBottomDepth?: boolean;
```

Property Value
boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### bevelBottomInset
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the inset size for the bottom bevel.

```typescript
bevelBottomInset?: boolean;
```

Property Value
boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### bevelBottomType
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a `BevelType` value that represents the bevel type for the bottom bevel.

```typescript
bevelBottomType?: boolean;
```

Property Value
boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### bevelTopDepth
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the depth of the top bevel.

```typescript
bevelTopDepth?: boolean;
```

Property Value
boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### bevelTopInset
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the inset size for the top bevel.

```typescript
bevelTopInset?: boolean;
```

Property Value
boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### bevelTopType
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a `BevelType` value that represents the bevel type for the top bevel.

```typescript
bevelTopType?: boolean;
```

Property Value
boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### contourColor
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `ColorFormat` object that represents color of the contour of a shape.

```typescript
contourColor?: Word.Interfaces.ColorFormatLoadOptions;
```

Property Value
[Word.Interfaces.ColorFormatLoadOptions](/en-us/javascript/api/word/word.interfaces.colorformatloadoptions)

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### contourWidth
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the width of the contour of a shape.

```typescript
contourWidth?: boolean;
```

Property Value
boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### depth
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the depth of the shape's extrusion.

```typescript
depth?: boolean;
```

Property Value
boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### extrusionColor
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `ColorFormat` object that represents the color of the shape's extrusion.

```typescript
extrusionColor?: Word.Interfaces.ColorFormatLoadOptions;
```

Property Value
[Word.Interfaces.ColorFormatLoadOptions](/en-us/javascript/api/word/word.interfaces.colorformatloadoptions)

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### extrusionColorType
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the extrusion color is based on the extruded shape's fill (the front face of the extrusion) and automatically changes when the shape's fill changes, or whether the extrusion color is independent of the shape's fill.

```typescript
extrusionColorType?: boolean;
```

Property Value
boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### fieldOfView
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the amount of perspective for a shape.

```typescript
fieldOfView?: boolean;
```

Property Value
boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### isPerspective
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies `true` if the extrusion appears in perspective â that is, if the walls of the extrusion narrow toward a vanishing point, `false` if the extrusion is a parallel, or orthographic, projection â that is, if the walls don't narrow toward a vanishing point.

```typescript
isPerspective?: boolean;
```

Property Value
boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### isVisible
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the specified object, or the formatting applied to it, is visible.

```typescript
isVisible?: boolean;
```

Property Value
boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### lightAngle
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the angle of the lighting.

```typescript
lightAngle?: boolean;
```

Property Value
boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### presetCamera
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `PresetCamera` value that represents the camera presets.

```typescript
presetCamera?: boolean;
```

Property Value
boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### presetExtrusionDirection
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the direction taken by the extrusion's sweep path leading away from the extruded shape (the front face of the extrusion).

```typescript
presetExtrusionDirection?: boolean;
```

Property Value
boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### presetLighting
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a `LightRigType` value that represents the lighting preset.

```typescript
presetLighting?: boolean;
```

Property Value
boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### presetLightingDirection
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the position of the light source relative to the extrusion.

```typescript
presetLightingDirection?: boolean;
```

Property Value
boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### presetLightingSoftness
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the intensity of the extrusion lighting.

```typescript
presetLightingSoftness?: boolean;
```

Property Value
boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### presetMaterial
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the extrusion surface material.

```typescript
presetMaterial?: boolean;
```

Property Value
boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### presetThreeDimensionalFormat
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the preset extrusion format.

```typescript
presetThreeDimensionalFormat?: boolean;
```

Property Value
boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### projectText
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether text on a shape rotates with shape. `true` rotates the text.

```typescript
projectText?: boolean;
```

Property Value
boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### rotationX
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the rotation of the extruded shape around the x-axis in degrees.

```typescript
rotationX?: boolean;
```

Property Value
boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### rotationY
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the rotation of the extruded shape around the y-axis in degrees.

```typescript
rotationY?: boolean;
```

Property Value
boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### rotationZ
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the z-axis rotation of the camera.

```typescript
rotationZ?: boolean;
```

Property Value
boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### z
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the position on the z-axis for the shape.

```typescript
z?: boolean;
```

Property Value
boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)