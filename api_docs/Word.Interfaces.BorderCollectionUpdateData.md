# Word.Interfaces.BorderCollectionUpdateData interface

Package: [word](/en-us/javascript/api/word)

An interface for updating data on the BorderCollection object, for use in borderCollection.set({ ... }).

## Properties

- insideBorderColor  
  Specifies the 24-bit color of the inside borders. Color is specified in '#RRGGBB' format or by using the color name.

- insideBorderType  
  Specifies the border type of the inside borders.

- insideBorderWidth  
  Specifies the width of the inside borders.

- items

- outsideBorderColor  
  Specifies the 24-bit color of the outside borders. Color is specified in '#RRGGBB' format or by using the color name.

- outsideBorderType  
  Specifies the border type of the outside borders.

- outsideBorderWidth  
  Specifies the width of the outside borders.

## Property Details

### insideBorderColor

Specifies the 24-bit color of the inside borders. Color is specified in '#RRGGBB' format or by using the color name.

```typescript
insideBorderColor?: string;
```

Property Value  
string

Remarks  
[API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### insideBorderType

Specifies the border type of the inside borders.

```typescript
insideBorderType?: Word.BorderType | "Mixed" | "None" | "Single" | "Double" | "Dotted" | "Dashed" | "DotDashed" | "Dot2Dashed" | "Triple" | "ThinThickSmall" | "ThickThinSmall" | "ThinThickThinSmall" | "ThinThickMed" | "ThickThinMed" | "ThinThickThinMed" | "ThinThickLarge" | "ThickThinLarge" | "ThinThickThinLarge" | "Wave" | "DoubleWave" | "DashedSmall" | "DashDotStroked" | "ThreeDEmboss" | "ThreeDEngrave";
```

Property Value  
[Word.BorderType](/en-us/javascript/api/word/word.bordertype) | "Mixed" | "None" | "Single" | "Double" | "Dotted" | "Dashed" | "DotDashed" | "Dot2Dashed" | "Triple" | "ThinThickSmall" | "ThickThinSmall" | "ThinThickThinSmall" | "ThinThickMed" | "ThickThinMed" | "ThinThickThinMed" | "ThinThickLarge" | "ThickThinLarge" | "ThinThickThinLarge" | "Wave" | "DoubleWave" | "DashedSmall" | "DashDotStroked" | "ThreeDEmboss" | "ThreeDEngrave"

Remarks  
[API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### insideBorderWidth

Specifies the width of the inside borders.

```typescript
insideBorderWidth?: Word.BorderWidth | "None" | "Pt025" | "Pt050" | "Pt075" | "Pt100" | "Pt150" | "Pt225" | "Pt300" | "Pt450" | "Pt600" | "Mixed";
```

Property Value  
[Word.BorderWidth](/en-us/javascript/api/word/word.borderwidth) | "None" | "Pt025" | "Pt050" | "Pt075" | "Pt100" | "Pt150" | "Pt225" | "Pt300" | "Pt450" | "Pt600" | "Mixed"

Remarks  
[API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### items

```typescript
items?: Word.Interfaces.BorderData[];
```

Property Value  
[Word.Interfaces.BorderData](/en-us/javascript/api/word/word.interfaces.borderdata)[]

---

### outsideBorderColor

Specifies the 24-bit color of the outside borders. Color is specified in '#RRGGBB' format or by using the color name.

```typescript
outsideBorderColor?: string;
```

Property Value  
string

Remarks  
[API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### outsideBorderType

Specifies the border type of the outside borders.

```typescript
outsideBorderType?: Word.BorderType | "Mixed" | "None" | "Single" | "Double" | "Dotted" | "Dashed" | "DotDashed" | "Dot2Dashed" | "Triple" | "ThinThickSmall" | "ThickThinSmall" | "ThinThickThinSmall" | "ThinThickMed" | "ThickThinMed" | "ThinThickThinMed" | "ThinThickLarge" | "ThickThinLarge" | "ThinThickThinLarge" | "Wave" | "DoubleWave" | "DashedSmall" | "DashDotStroked" | "ThreeDEmboss" | "ThreeDEngrave";
```

Property Value  
[Word.BorderType](/en-us/javascript/api/word/word.bordertype) | "Mixed" | "None" | "Single" | "Double" | "Dotted" | "Dashed" | "DotDashed" | "Dot2Dashed" | "Triple" | "ThinThickSmall" | "ThickThinSmall" | "ThinThickThinSmall" | "ThinThickMed" | "ThickThinMed" | "ThinThickThinMed" | "ThinThickLarge" | "ThickThinLarge" | "ThinThickThinLarge" | "Wave" | "DoubleWave" | "DashedSmall" | "DashDotStroked" | "ThreeDEmboss" | "ThreeDEngrave"

Remarks  
[API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### outsideBorderWidth

Specifies the width of the outside borders.

```typescript
outsideBorderWidth?: Word.BorderWidth | "None" | "Pt025" | "Pt050" | "Pt075" | "Pt100" | "Pt150" | "Pt225" | "Pt300" | "Pt450" | "Pt600" | "Mixed";
```

Property Value  
[Word.BorderWidth](/en-us/javascript/api/word/word.borderwidth) | "None" | "Pt025" | "Pt050" | "Pt075" | "Pt100" | "Pt150" | "Pt225" | "Pt300" | "Pt450" | "Pt600" | "Mixed"

Remarks  
[API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)