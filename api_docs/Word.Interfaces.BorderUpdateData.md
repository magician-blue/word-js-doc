# Word.Interfaces.BorderUpdateData interface

Package: [word](/en-us/javascript/api/word)

An interface for updating data on the Border object, for use in border.set({ ... }).

## Properties

- color — Specifies the color for the border. Color is specified in â#RRGGBBâ format or by using the color name.
- type — Specifies the border type for the border.
- visible — Specifies whether the border is visible.
- width — Specifies the width for the border.

## Property Details

### color

Specifies the color for the border. Color is specified in â#RRGGBBâ format or by using the color name.

```typescript
color?: string;
```

Property Value: string

Remarks  
[API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### type

Specifies the border type for the border.

```typescript
type?: Word.BorderType | "Mixed" | "None" | "Single" | "Double" | "Dotted" | "Dashed" | "DotDashed" | "Dot2Dashed" | "Triple" | "ThinThickSmall" | "ThickThinSmall" | "ThinThickThinSmall" | "ThinThickMed" | "ThickThinMed" | "ThinThickThinMed" | "ThinThickLarge" | "ThickThinLarge" | "ThinThickThinLarge" | "Wave" | "DoubleWave" | "DashedSmall" | "DashDotStroked" | "ThreeDEmboss" | "ThreeDEngrave";
```

Property Value: [Word.BorderType](/en-us/javascript/api/word/word.bordertype) | "Mixed" | "None" | "Single" | "Double" | "Dotted" | "Dashed" | "DotDashed" | "Dot2Dashed" | "Triple" | "ThinThickSmall" | "ThickThinSmall" | "ThinThickThinSmall" | "ThinThickMed" | "ThickThinMed" | "ThinThickThinMed" | "ThinThickLarge" | "ThickThinLarge" | "ThinThickThinLarge" | "Wave" | "DoubleWave" | "DashedSmall" | "DashDotStroked" | "ThreeDEmboss" | "ThreeDEngrave"

Remarks  
[API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### visible

Specifies whether the border is visible.

```typescript
visible?: boolean;
```

Property Value: boolean

Remarks  
[API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### width

Specifies the width for the border.

```typescript
width?: Word.BorderWidth | "None" | "Pt025" | "Pt050" | "Pt075" | "Pt100" | "Pt150" | "Pt225" | "Pt300" | "Pt450" | "Pt600" | "Mixed";
```

Property Value: [Word.BorderWidth](/en-us/javascript/api/word/word.borderwidth) | "None" | "Pt025" | "Pt050" | "Pt075" | "Pt100" | "Pt150" | "Pt225" | "Pt300" | "Pt450" | "Pt600" | "Mixed"

Remarks  
[API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)