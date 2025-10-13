# Word.Interfaces.TableBorderData interface

Package: [word](/en-us/javascript/api/word)

An interface describing the data returned by calling `tableBorder.toJSON()`.

## Properties

- color — Specifies the table border color.
- type — Specifies the type of the table border.
- width — Specifies the width, in points, of the table border. Not applicable to table border types that have fixed widths.

## Property Details

### color

Specifies the table border color.

```typescript
color?: string;
```

#### Property Value
string

#### Remarks
[API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### type

Specifies the type of the table border.

```typescript
type?: Word.BorderType | "Mixed" | "None" | "Single" | "Double" | "Dotted" | "Dashed" | "DotDashed" | "Dot2Dashed" | "Triple" | "ThinThickSmall" | "ThickThinSmall" | "ThinThickThinSmall" | "ThinThickMed" | "ThickThinMed" | "ThinThickThinMed" | "ThinThickLarge" | "ThickThinLarge" | "ThinThickThinLarge" | "Wave" | "DoubleWave" | "DashedSmall" | "DashDotStroked" | "ThreeDEmboss" | "ThreeDEngrave";
```

#### Property Value
[Word.BorderType](/en-us/javascript/api/word/word.bordertype) | "Mixed" | "None" | "Single" | "Double" | "Dotted" | "Dashed" | "DotDashed" | "Dot2Dashed" | "Triple" | "ThinThickSmall" | "ThickThinSmall" | "ThinThickThinSmall" | "ThinThickMed" | "ThickThinMed" | "ThinThickThinMed" | "ThinThickLarge" | "ThickThinLarge" | "ThinThickThinLarge" | "Wave" | "DoubleWave" | "DashedSmall" | "DashDotStroked" | "ThreeDEmboss" | "ThreeDEngrave"

#### Remarks
[API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### width

Specifies the width, in points, of the table border. Not applicable to table border types that have fixed widths.

```typescript
width?: number;
```

#### Property Value
number

#### Remarks
[API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)