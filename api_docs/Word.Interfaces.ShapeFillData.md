# Word.Interfaces.ShapeFillData interface

Package: [word](/en-us/javascript/api/word)

An interface describing the data returned by calling `shapeFill.toJSON()`.

## Properties

- [backgroundColor](#word-word-interfaces-shapefilldata-backgroundcolor-member)  
  Specifies the shape fill background color. You can provide the value in the '#RRGGBB' format or the color name.
- [foregroundColor](#word-word-interfaces-shapefilldata-foregroundcolor-member)  
  Specifies the shape fill foreground color. You can provide the value in the '#RRGGBB' format or the color name.
- [transparency](#word-word-interfaces-shapefilldata-transparency-member)  
  Specifies the transparency percentage of the fill as a value from 0.0 (opaque) through 1.0 (clear). Returns `null` if the shape type does not support transparency or the shape fill has inconsistent transparency, such as with a gradient fill type.
- [type](#word-word-interfaces-shapefilldata-type-member)  
  Returns the fill type of the shape. See `Word.ShapeFillType` for details.

## Property Details

### backgroundColor
<a id="word-word-interfaces-shapefilldata-backgroundcolor-member"></a>

Specifies the shape fill background color. You can provide the value in the '#RRGGBB' format or the color name.

```typescript
backgroundColor?: string;
```

#### Property value
string

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### foregroundColor
<a id="word-word-interfaces-shapefilldata-foregroundcolor-member"></a>

Specifies the shape fill foreground color. You can provide the value in the '#RRGGBB' format or the color name.

```typescript
foregroundColor?: string;
```

#### Property value
string

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### transparency
<a id="word-word-interfaces-shapefilldata-transparency-member"></a>

Specifies the transparency percentage of the fill as a value from 0.0 (opaque) through 1.0 (clear). Returns `null` if the shape type does not support transparency or the shape fill has inconsistent transparency, such as with a gradient fill type.

```typescript
transparency?: number;
```

#### Property value
number

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### type
<a id="word-word-interfaces-shapefilldata-type-member"></a>

Returns the fill type of the shape. See `Word.ShapeFillType` for details.

```typescript
type?: Word.ShapeFillType | "NoFill" | "Solid" | "Gradient" | "Pattern" | "Picture" | "Texture" | "Mixed";
```

#### Property value
[Word.ShapeFillType](/en-us/javascript/api/word/word.shapefilltype) | "NoFill" | "Solid" | "Gradient" | "Pattern" | "Picture" | "Texture" | "Mixed"

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)