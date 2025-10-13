# Word.Interfaces.LineFormatUpdateData interface

Package: [word](/en-us/javascript/api/word)

An interface for updating data on the LineFormat object, for use in `lineFormat.set({ ... })`.

## Properties

- backgroundColor  
  Gets a ColorFormat object that represents the background color for a patterned line.
- beginArrowheadLength  
  Specifies the length of the arrowhead at the beginning of the line.
- beginArrowheadStyle  
  Specifies the style of the arrowhead at the beginning of the line.
- beginArrowheadWidth  
  Specifies the width of the arrowhead at the beginning of the line.
- dashStyle  
  Specifies the dash style for the line.
- endArrowheadLength  
  Specifies the length of the arrowhead at the end of the line.
- endArrowheadStyle  
  Specifies the style of the arrowhead at the end of the line.
- endArrowheadWidth  
  Specifies the width of the arrowhead at the end of the line.
- foregroundColor  
  Gets a ColorFormat object that represents the foreground color for the line.
- insetPen  
  Specifies if to draw lines inside a shape.
- isVisible  
  Specifies if the object, or the formatting applied to it, is visible.
- pattern  
  Specifies the pattern applied to the line.
- style  
  Specifies the line format style.
- transparency  
  Specifies the degree of transparency of the line as a value between 0.0 (opaque) and 1.0 (clear).
- weight  
  Specifies the thickness of the line in points.

## Property Details

### backgroundColor

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets a `ColorFormat` object that represents the background color for a patterned line.

```typescript
backgroundColor?: Word.Interfaces.ColorFormatUpdateData;
```

Property Value  
[Word.Interfaces.ColorFormatUpdateData](/en-us/javascript/api/word/word.interfaces.colorformatupdatedata)

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### beginArrowheadLength

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the length of the arrowhead at the beginning of the line.

```typescript
beginArrowheadLength?: Word.ArrowheadLength | "Mixed" | "Short" | "Medium" | "Long";
```

Property Value  
[Word.ArrowheadLength](/en-us/javascript/api/word/word.arrowheadlength) | "Mixed" | "Short" | "Medium" | "Long"

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### beginArrowheadStyle

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the style of the arrowhead at the beginning of the line.

```typescript
beginArrowheadStyle?: Word.ArrowheadStyle | "Mixed" | "None" | "Triangle" | "Open" | "Stealth" | "Diamond" | "Oval";
```

Property Value  
[Word.ArrowheadStyle](/en-us/javascript/api/word/word.arrowheadstyle) | "Mixed" | "None" | "Triangle" | "Open" | "Stealth" | "Diamond" | "Oval"

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### beginArrowheadWidth

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the width of the arrowhead at the beginning of the line.

```typescript
beginArrowheadWidth?: Word.ArrowheadWidth | "Mixed" | "Narrow" | "Medium" | "Wide";
```

Property Value  
[Word.ArrowheadWidth](/en-us/javascript/api/word/word.arrowheadwidth) | "Mixed" | "Narrow" | "Medium" | "Wide"

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### dashStyle

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the dash style for the line.

```typescript
dashStyle?: Word.LineDashStyle | "Mixed" | "Solid" | "SquareDot" | "RoundDot" | "Dash" | "DashDot" | "DashDotDot" | "LongDash" | "LongDashDot" | "LongDashDotDot" | "SysDash" | "SysDot" | "SysDashDot";
```

Property Value  
[Word.LineDashStyle](/en-us/javascript/api/word/word.linedashstyle) | "Mixed" | "Solid" | "SquareDot" | "RoundDot" | "Dash" | "DashDot" | "DashDotDot" | "LongDash" | "LongDashDot" | "LongDashDotDot" | "SysDash" | "SysDot" | "SysDashDot"

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### endArrowheadLength

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the length of the arrowhead at the end of the line.

```typescript
endArrowheadLength?: Word.ArrowheadLength | "Mixed" | "Short" | "Medium" | "Long";
```

Property Value  
[Word.ArrowheadLength](/en-us/javascript/api/word/word.arrowheadlength) | "Mixed" | "Short" | "Medium" | "Long"

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### endArrowheadStyle

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the style of the arrowhead at the end of the line.

```typescript
endArrowheadStyle?: Word.ArrowheadStyle | "Mixed" | "None" | "Triangle" | "Open" | "Stealth" | "Diamond" | "Oval";
```

Property Value  
[Word.ArrowheadStyle](/en-us/javascript/api/word/word.arrowheadstyle) | "Mixed" | "None" | "Triangle" | "Open" | "Stealth" | "Diamond" | "Oval"

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### endArrowheadWidth

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the width of the arrowhead at the end of the line.

```typescript
endArrowheadWidth?: Word.ArrowheadWidth | "Mixed" | "Narrow" | "Medium" | "Wide";
```

Property Value  
[Word.ArrowheadWidth](/en-us/javascript/api/word/word.arrowheadwidth) | "Mixed" | "Narrow" | "Medium" | "Wide"

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### foregroundColor

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets a `ColorFormat` object that represents the foreground color for the line.

```typescript
foregroundColor?: Word.Interfaces.ColorFormatUpdateData;
```

Property Value  
[Word.Interfaces.ColorFormatUpdateData](/en-us/javascript/api/word/word.interfaces.colorformatupdatedata)

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### insetPen

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if to draw lines inside a shape.

```typescript
insetPen?: boolean;
```

Property Value  
boolean

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### isVisible

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the object, or the formatting applied to it, is visible.

```typescript
isVisible?: boolean;
```

Property Value  
boolean

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### pattern

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the pattern applied to the line.

```typescript
pattern?: Word.PatternType | "Mixed" | "Percent5" | "Percent10" | "Percent20" | "Percent25" | "Percent30" | "Percent40" | "Percent50" | "Percent60" | "Percent70" | "Percent75" | "Percent80" | "Percent90" | "DarkHorizontal" | "DarkVertical" | "DarkDownwardDiagonal" | "DarkUpwardDiagonal" | "SmallCheckerBoard" | "Trellis" | "LightHorizontal" | "LightVertical" | "LightDownwardDiagonal" | "LightUpwardDiagonal" | "SmallGrid" | "DottedDiamond" | "WideDownwardDiagonal" | "WideUpwardDiagonal" | "DashedUpwardDiagonal" | "DashedDownwardDiagonal" | "NarrowVertical" | "NarrowHorizontal" | "DashedVertical" | "DashedHorizontal" | "LargeConfetti" | "LargeGrid" | "HorizontalBrick" | "LargeCheckerBoard" | "SmallConfetti" | "ZigZag" | "SolidDiamond" | "DiagonalBrick" | "OutlinedDiamond" | "Plaid" | "Sphere" | "Weave" | "DottedGrid" | "Divot" | "Shingle" | "Wave" | "Horizontal" | "Vertical" | "Cross" | "DownwardDiagonal" | "UpwardDiagonal" | "DiagonalCross";
```

Property Value  
[Word.PatternType](/en-us/javascript/api/word/word.patterntype) | "Mixed" | "Percent5" | "Percent10" | "Percent20" | "Percent25" | "Percent30" | "Percent40" | "Percent50" | "Percent60" | "Percent70" | "Percent75" | "Percent80" | "Percent90" | "DarkHorizontal" | "DarkVertical" | "DarkDownwardDiagonal" | "DarkUpwardDiagonal" | "SmallCheckerBoard" | "Trellis" | "LightHorizontal" | "LightVertical" | "LightDownwardDiagonal" | "LightUpwardDiagonal" | "SmallGrid" | "DottedDiamond" | "WideDownwardDiagonal" | "WideUpwardDiagonal" | "DashedUpwardDiagonal" | "DashedDownwardDiagonal" | "NarrowVertical" | "NarrowHorizontal" | "DashedVertical" | "DashedHorizontal" | "LargeConfetti" | "LargeGrid" | "HorizontalBrick" | "LargeCheckerBoard" | "SmallConfetti" | "ZigZag" | "SolidDiamond" | "DiagonalBrick" | "OutlinedDiamond" | "Plaid" | "Sphere" | "Weave" | "DottedGrid" | "Divot" | "Shingle" | "Wave" | "Horizontal" | "Vertical" | "Cross" | "DownwardDiagonal" | "UpwardDiagonal" | "DiagonalCross"

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### style

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the line format style.

```typescript
style?: Word.LineFormatStyle | "Mixed" | "Single" | "ThinThin" | "ThinThick" | "ThickThin" | "ThickBetweenThin";
```

Property Value  
[Word.LineFormatStyle](/en-us/javascript/api/word/word.lineformatstyle) | "Mixed" | "Single" | "ThinThin" | "ThinThick" | "ThickThin" | "ThickBetweenThin"

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### transparency

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the degree of transparency of the line as a value between 0.0 (opaque) and 1.0 (clear).

```typescript
transparency?: number;
```

Property Value  
number

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### weight

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the thickness of the line in points.

```typescript
weight?: number;
```

Property Value  
number

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)