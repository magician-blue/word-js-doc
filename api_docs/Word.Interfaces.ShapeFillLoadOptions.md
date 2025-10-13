# Word.Interfaces.ShapeFillLoadOptions interface

Package: word

Represents the fill formatting of a shape object.

## Remarks

[ API set: WordApiDesktop 1.2 ]

## Properties

- $all: Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
- backgroundColor: Specifies the shape fill background color. You can provide the value in the '#RRGGBB' format or the color name.
- foregroundColor: Specifies the shape fill foreground color. You can provide the value in the '#RRGGBB' format or the color name.
- transparency: Specifies the transparency percentage of the fill as a value from 0.0 (opaque) through 1.0 (clear). Returns `null` if the shape type does not support transparency or the shape fill has inconsistent transparency, such as with a gradient fill type.
- type: Returns the fill type of the shape. See `Word.ShapeFillType` for details.

## Property Details

### $all

Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

Property Value: boolean

---

### backgroundColor

Specifies the shape fill background color. You can provide the value in the '#RRGGBB' format or the color name.

```typescript
backgroundColor?: boolean;
```

Property Value: boolean

Remarks

[ API set: WordApiDesktop 1.2 ]

---

### foregroundColor

Specifies the shape fill foreground color. You can provide the value in the '#RRGGBB' format or the color name.

```typescript
foregroundColor?: boolean;
```

Property Value: boolean

Remarks

[ API set: WordApiDesktop 1.2 ]

---

### transparency

Specifies the transparency percentage of the fill as a value from 0.0 (opaque) through 1.0 (clear). Returns `null` if the shape type does not support transparency or the shape fill has inconsistent transparency, such as with a gradient fill type.

```typescript
transparency?: boolean;
```

Property Value: boolean

Remarks

[ API set: WordApiDesktop 1.2 ]

---

### type

Returns the fill type of the shape. See `Word.ShapeFillType` for details.

```typescript
type?: boolean;
```

Property Value: boolean

Remarks

[ API set: WordApiDesktop 1.2 ]