# Word.Interfaces.ColorFormatLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents the color formatting of a shape or text in Word.

## Remarks

[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- $all: Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
- brightness: Specifies the brightness of a specified shape color. Valid values are from `-1` (darkest) to `1` (lightest), `0` represents neutral.
- objectThemeColor: Specifies the theme color for a color format.
- rgb: Specifies the red-green-blue (RGB) value of the specified color. You can provide the value in the '#RRGGBB' format.
- tintAndShade: Specifies the lightening or darkening of a specified shape's color. Valid values are from `-1` (darkest) to `1` (lightest), `0` represents neutral.
- type: Returns the shape color type.

## Property Details

### $all

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

Property Value: boolean

### brightness

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the brightness of a specified shape color. Valid values are from `-1` (darkest) to `1` (lightest), `0` represents neutral.

```typescript
brightness?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### objectThemeColor

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the theme color for a color format.

```typescript
objectThemeColor?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### rgb

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the red-green-blue (RGB) value of the specified color. You can provide the value in the '#RRGGBB' format.

```typescript
rgb?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### tintAndShade

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the lightening or darkening of a specified shape's color. Valid values are from `-1` (darkest) to `1` (lightest), `0` represents neutral.

```typescript
tintAndShade?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### type

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the shape color type.

```typescript
type?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)