# Word.Interfaces.ColorFormatUpdateData interface

Package: [word](/en-us/javascript/api/word)

An interface for updating data on the `ColorFormat` object, for use in `colorFormat.set({ ... })`.

## Properties

- brightness
  - Specifies the brightness of a specified shape color. Valid values are from -1 (darkest) to 1 (lightest), 0 represents neutral.
- objectThemeColor
  - Specifies the theme color for a color format.
- rgb
  - Specifies the red-green-blue (RGB) value of the specified color. You can provide the value in the '#RRGGBB' format.
- tintAndShade
  - Specifies the lightening or darkening of a specified shape's color. Valid values are from -1 (darkest) to 1 (lightest), 0 represents neutral.

## Property Details

### brightness

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the brightness of a specified shape color. Valid values are from -1 (darkest) to 1 (lightest), 0 represents neutral.

```typescript
brightness?: number;
```

Property value: number

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### objectThemeColor

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the theme color for a color format.

```typescript
objectThemeColor?: Word.ThemeColorIndex | "NotThemeColor" | "MainDark1" | "MainLight1" | "MainDark2" | "MainLight2" | "Accent1" | "Accent2" | "Accent3" | "Accent4" | "Accent5" | "Accent6" | "Hyperlink" | "HyperlinkFollowed" | "Background1" | "Text1" | "Background2" | "Text2";
```

Property value: [Word.ThemeColorIndex](/en-us/javascript/api/word/word.themecolorindex) | "NotThemeColor" | "MainDark1" | "MainLight1" | "MainDark2" | "MainLight2" | "Accent1" | "Accent2" | "Accent3" | "Accent4" | "Accent5" | "Accent6" | "Hyperlink" | "HyperlinkFollowed" | "Background1" | "Text1" | "Background2" | "Text2"

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### rgb

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the red-green-blue (RGB) value of the specified color. You can provide the value in the '#RRGGBB' format.

```typescript
rgb?: string;
```

Property value: string

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### tintAndShade

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the lightening or darkening of a specified shape's color. Valid values are from -1 (darkest) to 1 (lightest), 0 represents neutral.

```typescript
tintAndShade?: number;
```

Property value: number

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)