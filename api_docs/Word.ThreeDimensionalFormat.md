# Word.ThreeDimensionalFormat class

Package: word (https://learn.microsoft.com/en-us/javascript/api/word)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents a shape's three-dimensional formatting.

Extends: OfficeExtension.ClientObject (https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientobject)

Remarks
- API set: WordApi BETA (PREVIEW ONLY) — https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

## Properties

- bevelBottomDepth — Specifies the depth of the bottom bevel.
- bevelBottomInset — Specifies the inset size for the bottom bevel.
- bevelBottomType — Specifies a BevelType value that represents the bevel type for the bottom bevel.
- bevelTopDepth — Specifies the depth of the top bevel.
- bevelTopInset — Specifies the inset size for the top bevel.
- bevelTopType — Specifies a BevelType value that represents the bevel type for the top bevel.
- context — The request context associated with the object. This connects the add-in's process to the Office host application's process.
- contourColor — Returns a ColorFormat object that represents color of the contour of a shape.
- contourWidth — Specifies the width of the contour of a shape.
- depth — Specifies the depth of the shape's extrusion.
- extrusionColor — Returns a ColorFormat object that represents the color of the shape's extrusion.
- extrusionColorType — Specifies whether the extrusion color is based on the extruded shape's fill (the front face of the extrusion) and automatically changes when the shape's fill changes, or whether the extrusion color is independent of the shape's fill.
- fieldOfView — Specifies the amount of perspective for a shape.
- isPerspective — Specifies true if the extrusion appears in perspective — that is, if the walls of the extrusion narrow toward a vanishing point, false if the extrusion is a parallel, or orthographic, projection — that is, if the walls don't narrow toward a vanishing point.
- isVisible — Specifies if the specified object, or the formatting applied to it, is visible.
- lightAngle — Specifies the angle of the lighting.
- presetCamera — Returns a PresetCamera value that represents the camera presets.
- presetExtrusionDirection — Returns the direction taken by the extrusion's sweep path leading away from the extruded shape (the front face of the extrusion).
- presetLighting — Specifies a LightRigType value that represents the lighting preset.
- presetLightingDirection — Specifies the position of the light source relative to the extrusion.
- presetLightingSoftness — Specifies the intensity of the extrusion lighting.
- presetMaterial — Specifies the extrusion surface material.
- presetThreeDimensionalFormat — Returns the preset extrusion format.
- projectText — Specifies whether text on a shape rotates with shape. true rotates the text.
- rotationX — Specifies the rotation of the extruded shape around the x-axis in degrees.
- rotationY — Specifies the rotation of the extruded shape around the y-axis in degrees.
- rotationZ — Specifies the z-axis rotation of the camera.
- z — Specifies the position on the z-axis for the shape.

## Methods

- incrementRotationHorizontal(increment) — Horizontally rotates a shape on the x-axis. The number of degrees to rotate.
- incrementRotationVertical(increment) — Vertically rotates a shape on the y-axis. The number of degrees to rotate.
- incrementRotationX(increment) — Changes the rotation around the x-axis. The number of degrees to rotate.
- incrementRotationY(increment) — Changes the rotation around the y-axis. The number of degrees to rotate.
- incrementRotationZ(increment) — Rotates a shape on the z-axis. The number of degrees to rotate.
- load(options) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNames) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNamesAndPaths) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- resetRotation() — Resets the extrusion rotation around the x-axis, y-axis, and z-axis to 0.
- set(properties, options) — Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
- set(properties) — Sets multiple properties on the object at the same time, based on an existing loaded object.
- setExtrusionDirection(presetExtrusionDirection) — Sets the direction of the extrusion's sweep path.
- setExtrusionDirection(presetExtrusionDirection) — Sets the direction of the extrusion's sweep path.
- setPresetCamera(presetCamera) — Sets the camera preset for the shape. The preset camera type.
- setPresetCamera(presetCamera) — Sets the camera preset for the shape. The preset camera type.
- setThreeDimensionalFormat(presetThreeDimensionalFormat) — Sets the preset extrusion format. The preset format.
- setThreeDimensionalFormat(presetThreeDimensionalFormat) — Sets the preset extrusion format. The preset format.
- toJSON() — Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.ThreeDimensionalFormat object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.ThreeDimensionalFormatData) that contains shallow copies of any loaded child properties from the original object.
- track() — Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
- untrack() — Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

---

## Property Details

### bevelBottomDepth

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the depth of the bottom bevel.

```typescript
bevelBottomDepth: number;
```

Property Value
- number

Remarks
- API set: WordApi BETA (PREVIEW ONLY) — https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### bevelBottomInset

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the inset size for the bottom bevel.

```typescript
bevelBottomInset: number;
```

Property Value
- number

Remarks
- API set: WordApi BETA (PREVIEW ONLY)

### bevelBottomType

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a BevelType value that represents the bevel type for the bottom bevel.

```typescript
bevelBottomType: Word.BevelType | "mixed" | "none" | "relaxedInset" | "circle" | "slope" | "cross" | "angle" | "softRound" | "convex" | "coolSlant" | "divot" | "riblet" | "hardEdge" | "artDeco";
```

Property Value
- Word.BevelType (https://learn.microsoft.com/en-us/javascript/api/word/word.beveltype) | "mixed" | "none" | "relaxedInset" | "circle" | "slope" | "cross" | "angle" | "softRound" | "convex" | "coolSlant" | "divot" | "riblet" | "hardEdge" | "artDeco"

Remarks
- API set: WordApi BETA (PREVIEW ONLY)

### bevelTopDepth

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the depth of the top bevel.

```typescript
bevelTopDepth: number;
```

Property Value
- number

Remarks
- API set: WordApi BETA (PREVIEW ONLY)

### bevelTopInset

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the inset size for the top bevel.

```typescript
bevelTopInset: number;
```

Property Value
- number

Remarks
- API set: WordApi BETA (PREVIEW ONLY)

### bevelTopType

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a BevelType value that represents the bevel type for the top bevel.

```typescript
bevelTopType: Word.BevelType | "mixed" | "none" | "relaxedInset" | "circle" | "slope" | "cross" | "angle" | "softRound" | "convex" | "coolSlant" | "divot" | "riblet" | "hardEdge" | "artDeco";
```

Property Value
- Word.BevelType (https://learn.microsoft.com/en-us/javascript/api/word/word.beveltype) | "mixed" | "none" | "relaxedInset" | "circle" | "slope" | "cross" | "angle" | "softRound" | "convex" | "coolSlant" | "divot" | "riblet" | "hardEdge" | "artDeco"

Remarks
- API set: WordApi BETA (PREVIEW ONLY)

### context

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value
- Word.RequestContext (https://learn.microsoft.com/en-us/javascript/api/word/word.requestcontext)

### contourColor

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a ColorFormat object that represents color of the contour of a shape.

```typescript
readonly contourColor: Word.ColorFormat;
```

Property Value
- Word.ColorFormat (https://learn.microsoft.com/en-us/javascript/api/word/word.colorformat)

Remarks
- API set: WordApi BETA (PREVIEW ONLY)

### contourWidth

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the width of the contour of a shape.

```typescript
contourWidth: number;
```

Property Value
- number

Remarks
- API set: WordApi BETA (PREVIEW ONLY)

### depth

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the depth of the shape's extrusion.

```typescript
depth: number;
```

Property Value
- number

Remarks
- API set: WordApi BETA (PREVIEW ONLY)

### extrusionColor

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a ColorFormat object that represents the color of the shape's extrusion.

```typescript
readonly extrusionColor: Word.ColorFormat;
```

Property Value
- Word.ColorFormat (https://learn.microsoft.com/en-us/javascript/api/word/word.colorformat)

Remarks
- API set: WordApi BETA (PREVIEW ONLY)

### extrusionColorType

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the extrusion color is based on the extruded shape's fill (the front face of the extrusion) and automatically changes when the shape's fill changes, or whether the extrusion color is independent of the shape's fill.

```typescript
extrusionColorType: Word.ExtrusionColorType | "mixed" | "automatic" | "custom";
```

Property Value
- Word.ExtrusionColorType (https://learn.microsoft.com/en-us/javascript/api/word/word.extrusioncolortype) | "mixed" | "automatic" | "custom"

Remarks
- API set: WordApi BETA (PREVIEW ONLY)

### fieldOfView

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the amount of perspective for a shape.

```typescript
fieldOfView: number;
```

Property Value
- number

Remarks
- API set: WordApi BETA (PREVIEW ONLY)

### isPerspective

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies true if the extrusion appears in perspective — that is, if the walls of the extrusion narrow toward a vanishing point, false if the extrusion is a parallel, or orthographic, projection — that is, if the walls don't narrow toward a vanishing point.

```typescript
isPerspective: boolean;
```

Property Value
- boolean

Remarks
- API set: WordApi BETA (PREVIEW ONLY)

### isVisible

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the specified object, or the formatting applied to it, is visible.

```typescript
isVisible: boolean;
```

Property Value
- boolean

Remarks
- API set: WordApi BETA (PREVIEW ONLY)

### lightAngle

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the angle of the lighting.

```typescript
lightAngle: number;
```

Property Value
- number

Remarks
- API set: WordApi BETA (PREVIEW ONLY)

### presetCamera

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a PresetCamera value that represents the camera presets.

```typescript
readonly presetCamera: Word.PresetCamera | "Mixed" | "LegacyObliqueTopLeft" | "LegacyObliqueTop" | "LegacyObliqueTopRight" | "LegacyObliqueLeft" | "LegacyObliqueFront" | "LegacyObliqueRight" | "LegacyObliqueBottomLeft" | "LegacyObliqueBottom" | "LegacyObliqueBottomRight" | "LegacyPerspectiveTopLeft" | "LegacyPerspectiveTop" | "LegacyPerspectiveTopRight" | "LegacyPerspectiveLeft" | "LegacyPerspectiveFront" | "LegacyPerspectiveRight" | "LegacyPerspectiveBottomLeft" | "LegacyPerspectiveBottom" | "LegacyPerspectiveBottomRight" | "OrthographicFront" | "IsometricTopUp" | "IsometricTopDown" | "IsometricBottomUp" | "IsometricBottomDown" | "IsometricLeftUp" | "IsometricLeftDown" | "IsometricRightUp" | "IsometricRightDown" | "IsometricOffAxis1Left" | "IsometricOffAxis1Right" | "IsometricOffAxis1Top" | "IsometricOffAxis2Left" | "IsometricOffAxis2Right" | "IsometricOffAxis2Top" | "IsometricOffAxis3Left" | "IsometricOffAxis3Right" | "IsometricOffAxis3Bottom" | "IsometricOffAxis4Left" | "IsometricOffAxis4Right" | "IsometricOffAxis4Bottom" | "ObliqueTopLeft" | "ObliqueTop" | "ObliqueTopRight" | "ObliqueLeft" | "ObliqueRight" | "ObliqueBottomLeft" | "ObliqueBottom" | "ObliqueBottomRight" | "PerspectiveFront" | "PerspectiveLeft" | "PerspectiveRight" | "PerspectiveAbove" | "PerspectiveBelow" | "PerspectiveAboveLeftFacing" | "PerspectiveAboveRightFacing" | "PerspectiveContrastingLeftFacing" | "PerspectiveContrastingRightFacing" | "PerspectiveHeroicLeftFacing" | "PerspectiveHeroicRightFacing" | "PerspectiveHeroicExtremeLeftFacing" | "PerspectiveHeroicExtremeRightFacing" | "PerspectiveRelaxed" | "PerspectiveRelaxedModerately";
```

Property Value
- Word.PresetCamera (https://learn.microsoft.com/en-us/javascript/api/word/word.presetcamera) | "Mixed" | "LegacyObliqueTopLeft" | "LegacyObliqueTop" | "LegacyObliqueTopRight" | "LegacyObliqueLeft" | "LegacyObliqueFront" | "LegacyObliqueRight" | "LegacyObliqueBottomLeft" | "LegacyObliqueBottom" | "LegacyObliqueBottomRight" | "LegacyPerspectiveTopLeft" | "LegacyPerspectiveTop" | "LegacyPerspectiveTopRight" | "LegacyPerspectiveLeft" | "LegacyPerspectiveFront" | "LegacyPerspectiveRight" | "LegacyPerspectiveBottomLeft" | "LegacyPerspectiveBottom" | "LegacyPerspectiveBottomRight" | "OrthographicFront" | "IsometricTopUp" | "IsometricTopDown" | "IsometricBottomUp" | "IsometricBottomDown" | "IsometricLeftUp" | "IsometricLeftDown" | "IsometricRightUp" | "IsometricRightDown" | "IsometricOffAxis1Left" | "IsometricOffAxis1Right" | "IsometricOffAxis1Top" | "IsometricOffAxis2Left" | "IsometricOffAxis2Right" | "IsometricOffAxis2Top" | "IsometricOffAxis3Left" | "IsometricOffAxis3Right" | "IsometricOffAxis3Bottom" | "IsometricOffAxis4Left" | "IsometricOffAxis4Right" | "IsometricOffAxis4Bottom" | "ObliqueTopLeft" | "ObliqueTop" | "ObliqueTopRight" | "ObliqueLeft" | "ObliqueRight" | "ObliqueBottomLeft" | "ObliqueBottom" | "ObliqueBottomRight" | "PerspectiveFront" | "PerspectiveLeft" | "PerspectiveRight" | "PerspectiveAbove" | "PerspectiveBelow" | "PerspectiveAboveLeftFacing" | "PerspectiveAboveRightFacing" | "PerspectiveContrastingLeftFacing" | "PerspectiveContrastingRightFacing" | "PerspectiveHeroicLeftFacing" | "PerspectiveHeroicRightFacing" | "PerspectiveHeroicExtremeLeftFacing" | "PerspectiveHeroicExtremeRightFacing" | "PerspectiveRelaxed" | "PerspectiveRelaxedModerately"

Remarks
- API set: WordApi BETA (PREVIEW ONLY)

### presetExtrusionDirection

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the direction taken by the extrusion's sweep path leading away from the extruded shape (the front face of the extrusion).

```typescript
readonly presetExtrusionDirection: Word.PresetExtrusionDirection | "Mixed" | "BottomRight" | "Bottom" | "BottomLeft" | "Right" | "None" | "Left" | "TopRight" | "Top" | "TopLeft";
```

Property Value
- Word.PresetExtrusionDirection (https://learn.microsoft.com/en-us/javascript/api/word/word.presetextrusiondirection) | "Mixed" | "BottomRight" | "Bottom" | "BottomLeft" | "Right" | "None" | "Left" | "TopRight" | "Top" | "TopLeft"

Remarks
- API set: WordApi BETA (PREVIEW ONLY)

### presetLighting

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a LightRigType value that represents the lighting preset.

```typescript
presetLighting: Word.LightRigType | "Mixed" | "LegacyFlat1" | "LegacyFlat2" | "LegacyFlat3" | "LegacyFlat4" | "LegacyNormal1" | "LegacyNormal2" | "LegacyNormal3" | "LegacyNormal4" | "LegacyHarsh1" | "LegacyHarsh2" | "LegacyHarsh3" | "LegacyHarsh4" | "ThreePoint" | "Balanced" | "Soft" | "Harsh" | "Flood" | "Contrasting" | "Morning" | "Sunrise" | "Sunset" | "Chilly" | "Freezing" | "Flat" | "TwoPoint" | "Glow" | "BrightRoom";
```

Property Value
- Word.LightRigType (https://learn.microsoft.com/en-us/javascript/api/word/word.lightrigtype) | "Mixed" | "LegacyFlat1" | "LegacyFlat2" | "LegacyFlat3" | "LegacyFlat4" | "LegacyNormal1" | "LegacyNormal2" | "LegacyNormal3" | "LegacyNormal4" | "LegacyHarsh1" | "LegacyHarsh2" | "LegacyHarsh3" | "LegacyHarsh4" | "ThreePoint" | "Balanced" | "Soft" | "Harsh" | "Flood" | "Contrasting" | "Morning" | "Sunrise" | "Sunset" | "Chilly" | "Freezing" | "Flat" | "TwoPoint" | "Glow" | "BrightRoom"

Remarks
- API set: WordApi BETA (PREVIEW ONLY)

### presetLightingDirection

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the position of the light source relative to the extrusion.

```typescript
presetLightingDirection: Word.PresetLightingDirection | "Mixed" | "TopLeft" | "Top" | "TopRight" | "Left" | "None" | "Right" | "BottomLeft" | "Bottom" | "BottomRight";
```

Property Value
- Word.PresetLightingDirection (https://learn.microsoft.com/en-us/javascript/api/word/word.presetlightingdirection) | "Mixed" | "TopLeft" | "Top" | "TopRight" | "Left" | "None" | "Right" | "BottomLeft" | "Bottom" | "BottomRight"

Remarks
- API set: WordApi BETA (PREVIEW ONLY)

### presetLightingSoftness

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the intensity of the extrusion lighting.

```typescript
presetLightingSoftness: Word.PresetLightingSoftness | "Mixed" | "Dim" | "Normal" | "Bright";
```

Property Value
- Word.PresetLightingSoftness (https://learn.microsoft.com/en-us/javascript/api/word/word.presetlightingsoftness) | "Mixed" | "Dim" | "Normal" | "Bright"

Remarks
- API set: WordApi BETA (PREVIEW ONLY)

### presetMaterial

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the extrusion surface material.

```typescript
presetMaterial: Word.PresetMaterial | "Mixed" | "Matte" | "Plastic" | "Metal" | "WireFrame" | "Matte2" | "Plastic2" | "Metal2" | "WarmMatte" | "TranslucentPowder" | "Powder" | "DarkEdge" | "SoftEdge" | "Clear" | "Flat" | "SoftMetal";
```

Property Value
- Word.PresetMaterial (https://learn.microsoft.com/en-us/javascript/api/word/word.presetmaterial) | "Mixed" | "Matte" | "Plastic" | "Metal" | "WireFrame" | "Matte2" | "Plastic2" | "Metal2" | "WarmMatte" | "TranslucentPowder" | "Powder" | "DarkEdge" | "SoftEdge" | "Clear" | "Flat" | "SoftMetal"

Remarks
- API set: WordApi BETA (PREVIEW ONLY)

### presetThreeDimensionalFormat

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the preset extrusion format.

```typescript
readonly presetThreeDimensionalFormat: Word.PresetThreeDimensionalFormat | "Mixed" | "Format1" | "Format2" | "Format3" | "Format4" | "Format5" | "Format6" | "Format7" | "Format8" | "Format9" | "Format10" | "Format11" | "Format12" | "Format13" | "Format14" | "Format15" | "Format16" | "Format17" | "Format18" | "Format19" | "Format20";
```

Property Value
- Word.PresetThreeDimensionalFormat (https://learn.microsoft.com/en-us/javascript/api/word/word.presetthreedimensionalformat) | "Mixed" | "Format1" | "Format2" | "Format3" | "Format4" | "Format5" | "Format6" | "Format7" | "Format8" | "Format9" | "Format10" | "Format11" | "Format12" | "Format13" | "Format14" | "Format15" | "Format16" | "Format17" | "Format18" | "Format19" | "Format20"

Remarks
- API set: WordApi BETA (PREVIEW ONLY)

### projectText

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether text on a shape rotates with shape. true rotates the text.

```typescript
projectText: boolean;
```

Property Value
- boolean

Remarks
- API set: WordApi BETA (PREVIEW ONLY)

### rotationX

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the rotation of the extruded shape around the x-axis in degrees.

```typescript
rotationX: number;
```

Property Value
- number

Remarks
- API set: WordApi BETA (PREVIEW ONLY)

### rotationY

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the rotation of the extruded shape around the y-axis in degrees.

```typescript
rotationY: number;
```

Property Value
- number

Remarks
- API set: WordApi BETA (PREVIEW ONLY)

### rotationZ

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the z-axis rotation of the camera.

```typescript
rotationZ: number;
```

Property Value
- number

Remarks
- API set: WordApi BETA (PREVIEW ONLY)

### z

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the position on the z-axis for the shape.

```typescript
z: number;
```

Property Value
- number

Remarks
- API set: WordApi BETA (PREVIEW ONLY)

---

## Method Details

### incrementRotationHorizontal(increment)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Horizontally rotates a shape on the x-axis. The number of degrees to rotate.

```typescript
incrementRotationHorizontal(increment: number): void;
```

Parameters
- increment: number

Returns
- void

Remarks
- API set: WordApi BETA (PREVIEW ONLY)

### incrementRotationVertical(increment)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Vertically rotates a shape on the y-axis. The number of degrees to rotate.

```typescript
incrementRotationVertical(increment: number): void;
```

Parameters
- increment: number

Returns
- void

Remarks
- API set: WordApi BETA (PREVIEW ONLY)

### incrementRotationX(increment)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Changes the rotation around the x-axis. The number of degrees to rotate.

```typescript
incrementRotationX(increment: number): void;
```

Parameters
- increment: number

Returns
- void

Remarks
- API set: WordApi BETA (PREVIEW ONLY)

### incrementRotationY(increment)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Changes the rotation around the y-axis. The number of degrees to rotate.

```typescript
incrementRotationY(increment: number): void;
```

Parameters
- increment: number

Returns
- void

Remarks
- API set: WordApi BETA (PREVIEW ONLY)

### incrementRotationZ(increment)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Rotates a shape on the z-axis. The number of degrees to rotate.

```typescript
incrementRotationZ(increment: number): void;
```

Parameters
- increment: number

Returns
- void

Remarks
- API set: WordApi BETA (PREVIEW ONLY)

### load(options)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(options?: Word.Interfaces.ThreeDimensionalFormatLoadOptions): Word.ThreeDimensionalFormat;
```

Parameters
- options: Word.Interfaces.ThreeDimensionalFormatLoadOptions (https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.threedimensionalformatloadoptions)  
  Provides options for which properties of the object to load.

Returns
- Word.ThreeDimensionalFormat (https://learn.microsoft.com/en-us/javascript/api/word/word.threedimensionalformat)

### load(propertyNames)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.ThreeDimensionalFormat;
```

Parameters
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns
- Word.ThreeDimensionalFormat (https://learn.microsoft.com/en-us/javascript/api/word/word.threedimensionalformat)

### load(propertyNamesAndPaths)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
    select?: string;
    expand?: string;
}): Word.ThreeDimensionalFormat;
```

Parameters
- propertyNamesAndPaths: { select?: string; expand?: string; }  
  propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

Returns
- Word.ThreeDimensionalFormat (https://learn.microsoft.com/en-us/javascript/api/word/word.threedimensionalformat)

### resetRotation()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Resets the extrusion rotation around the x-axis, y-axis, and z-axis to 0.

```typescript
resetRotation(): void;
```

Returns
- void

Remarks
- API set: WordApi BETA (PREVIEW ONLY)

### set(properties, options)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.ThreeDimensionalFormatUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

Parameters
- properties: Word.Interfaces.ThreeDimensionalFormatUpdateData (https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.threedimensionalformatupdatedata)  
  A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
- options: OfficeExtension.UpdateOptions (https://learn.microsoft.com/en-us/javascript/api/office/officeextension.updateoptions)  
  Provides an option to suppress errors if the properties object tries to set any read-only properties.

Returns
- void

### set(properties)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.ThreeDimensionalFormat): void;
```

Parameters
- properties: Word.ThreeDimensionalFormat (https://learn.microsoft.com/en-us/javascript/api/word/word.threedimensionalformat)

Returns
- void

### setExtrusionDirection(presetExtrusionDirection)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets the direction of the extrusion's sweep path.

```typescript
setExtrusionDirection(presetExtrusionDirection: Word.PresetExtrusionDirection): void;
```

Parameters
- presetExtrusionDirection: Word.PresetExtrusionDirection (https://learn.microsoft.com/en-us/javascript/api/word/word.presetextrusiondirection)  
  The preset direction.

Returns
- void

Remarks
- API set: WordApi BETA (PREVIEW ONLY)

### setExtrusionDirection(presetExtrusionDirection)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets the direction of the extrusion's sweep path.

```typescript
setExtrusionDirection(presetExtrusionDirection: "Mixed" | "BottomRight" | "Bottom" | "BottomLeft" | "Right" | "None" | "Left" | "TopRight" | "Top" | "TopLeft"): void;
```

Parameters
- presetExtrusionDirection: "Mixed" | "BottomRight" | "Bottom" | "BottomLeft" | "Right" | "None" | "Left" | "TopRight" | "Top" | "TopLeft"  
  The preset direction.

Returns
- void

Remarks
- API set: WordApi BETA (PREVIEW ONLY)

### setPresetCamera(presetCamera)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets the camera preset for the shape. The preset camera type.

```typescript
setPresetCamera(presetCamera: Word.PresetCamera): void;
```

Parameters
- presetCamera: Word.PresetCamera (https://learn.microsoft.com/en-us/javascript/api/word/word.presetcamera)

Returns
- void

Remarks
- API set: WordApi BETA (PREVIEW ONLY)

### setPresetCamera(presetCamera)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets the camera preset for the shape. The preset camera type.

```typescript
setPresetCamera(presetCamera: "Mixed" | "LegacyObliqueTopLeft" | "LegacyObliqueTop" | "LegacyObliqueTopRight" | "LegacyObliqueLeft" | "LegacyObliqueFront" | "LegacyObliqueRight" | "LegacyObliqueBottomLeft" | "LegacyObliqueBottom" | "LegacyObliqueBottomRight" | "LegacyPerspectiveTopLeft" | "LegacyPerspectiveTop" | "LegacyPerspectiveTopRight" | "LegacyPerspectiveLeft" | "LegacyPerspectiveFront" | "LegacyPerspectiveRight" | "LegacyPerspectiveBottomLeft" | "LegacyPerspectiveBottom" | "LegacyPerspectiveBottomRight" | "OrthographicFront" | "IsometricTopUp" | "IsometricTopDown" | "IsometricBottomUp" | "IsometricBottomDown" | "IsometricLeftUp" | "IsometricLeftDown" | "IsometricRightUp" | "IsometricRightDown" | "IsometricOffAxis1Left" | "IsometricOffAxis1Right" | "IsometricOffAxis1Top" | "IsometricOffAxis2Left" | "IsometricOffAxis2Right" | "IsometricOffAxis2Top" | "IsometricOffAxis3Left" | "IsometricOffAxis3Right" | "IsometricOffAxis3Bottom" | "IsometricOffAxis4Left" | "IsometricOffAxis4Right" | "IsometricOffAxis4Bottom" | "ObliqueTopLeft" | "ObliqueTop" | "ObliqueTopRight" | "ObliqueLeft" | "ObliqueRight" | "ObliqueBottomLeft" | "ObliqueBottom" | "ObliqueBottomRight" | "PerspectiveFront" | "PerspectiveLeft" | "PerspectiveRight" | "PerspectiveAbove" | "PerspectiveBelow" | "PerspectiveAboveLeftFacing" | "PerspectiveAboveRightFacing" | "PerspectiveContrastingLeftFacing" | "PerspectiveContrastingRightFacing" | "PerspectiveHeroicLeftFacing" | "PerspectiveHeroicRightFacing" | "PerspectiveHeroicExtremeLeftFacing" | "PerspectiveHeroicExtremeRightFacing" | "PerspectiveRelaxed" | "PerspectiveRelaxedModerately"): void;
```

Parameters
- presetCamera: "Mixed" | "LegacyObliqueTopLeft" | "LegacyObliqueTop" | "LegacyObliqueTopRight" | "LegacyObliqueLeft" | "LegacyObliqueFront" | "LegacyObliqueRight" | "LegacyObliqueBottomLeft" | "LegacyObliqueBottom" | "LegacyObliqueBottomRight" | "LegacyPerspectiveTopLeft" | "LegacyPerspectiveTop" | "LegacyPerspectiveTopRight" | "LegacyPerspectiveLeft" | "LegacyPerspectiveFront" | "LegacyPerspectiveRight" | "LegacyPerspectiveBottomLeft" | "LegacyPerspectiveBottom" | "LegacyPerspectiveBottomRight" | "OrthographicFront" | "IsometricTopUp" | "IsometricTopDown" | "IsometricBottomUp" | "IsometricBottomDown" | "IsometricLeftUp" | "IsometricLeftDown" | "IsometricRightUp" | "IsometricRightDown" | "IsometricOffAxis1Left" | "IsometricOffAxis1Right" | "IsometricOffAxis1Top" | "IsometricOffAxis2Left" | "IsometricOffAxis2Right" | "IsometricOffAxis2Top" | "IsometricOffAxis3Left" | "IsometricOffAxis3Right" | "IsometricOffAxis3Bottom" | "IsometricOffAxis4Left" | "IsometricOffAxis4Right" | "IsometricOffAxis4Bottom" | "ObliqueTopLeft" | "ObliqueTop" | "ObliqueTopRight" | "ObliqueLeft" | "ObliqueRight" | "ObliqueBottomLeft" | "ObliqueBottom" | "ObliqueBottomRight" | "PerspectiveFront" | "PerspectiveLeft" | "PerspectiveRight" | "PerspectiveAbove" | "PerspectiveBelow" | "PerspectiveAboveLeftFacing" | "PerspectiveAboveRightFacing" | "PerspectiveContrastingLeftFacing" | "PerspectiveContrastingRightFacing" | "PerspectiveHeroicLeftFacing" | "PerspectiveHeroicRightFacing" | "PerspectiveHeroicExtremeLeftFacing" | "PerspectiveHeroicExtremeRightFacing" | "PerspectiveRelaxed" | "PerspectiveRelaxedModerately"

Returns
- void

Remarks
- API set: WordApi BETA (PREVIEW ONLY)

### setThreeDimensionalFormat(presetThreeDimensionalFormat)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets the preset extrusion format. The preset format.

```typescript
setThreeDimensionalFormat(presetThreeDimensionalFormat: Word.PresetThreeDimensionalFormat): void;
```

Parameters
- presetThreeDimensionalFormat: Word.PresetThreeDimensionalFormat (https://learn.microsoft.com/en-us/javascript/api/word/word.presetthreedimensionalformat)

Returns
- void

Remarks
- API set: WordApi BETA (PREVIEW ONLY)

### setThreeDimensionalFormat(presetThreeDimensionalFormat)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets the preset extrusion format. The preset format.

```typescript
setThreeDimensionalFormat(presetThreeDimensionalFormat: "Mixed" | "Format1" | "Format2" | "Format3" | "Format4" | "Format5" | "Format6" | "Format7" | "Format8" | "Format9" | "Format10" | "Format11" | "Format12" | "Format13" | "Format14" | "Format15" | "Format16" | "Format17" | "Format18" | "Format19" | "Format20"): void;
```

Parameters
- presetThreeDimensionalFormat: "Mixed" | "Format1" | "Format2" | "Format3" | "Format4" | "Format5" | "Format6" | "Format7" | "Format8" | "Format9" | "Format10" | "Format11" | "Format12" | "Format13" | "Format14" | "Format15" | "Format16" | "Format17" | "Format18" | "Format19" | "Format20"

Returns
- void

Remarks
- API set: WordApi BETA (PREVIEW ONLY)

### toJSON()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.ThreeDimensionalFormat object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.ThreeDimensionalFormatData) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.ThreeDimensionalFormatData;
```

Returns
- Word.Interfaces.ThreeDimensionalFormatData (https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.threedimensionalformatdata)

### track()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject) (https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.ThreeDimensionalFormat;
```

Returns
- Word.ThreeDimensionalFormat (https://learn.microsoft.com/en-us/javascript/api/word/word.threedimensionalformat)

### untrack()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject) (https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.ThreeDimensionalFormat;
```

Returns
- Word.ThreeDimensionalFormat (https://learn.microsoft.com/en-us/javascript/api/word/word.threedimensionalformat)