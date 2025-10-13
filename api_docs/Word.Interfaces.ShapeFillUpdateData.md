# Word.Interfaces.ShapeFillUpdateData interface

Package: [word](https://learn.microsoft.com/en-us/javascript/api/word)

An interface for updating data on the `ShapeFill` object, for use in `shapeFill.set({ ... })`.

## Properties

- backgroundColor  
  Specifies the shape fill background color. You can provide the value in the '#RRGGBB' format or the color name.
- foregroundColor  
  Specifies the shape fill foreground color. You can provide the value in the '#RRGGBB' format or the color name.
- transparency  
  Specifies the transparency percentage of the fill as a value from 0.0 (opaque) through 1.0 (clear). Returns `null` if the shape type does not support transparency or the shape fill has inconsistent transparency, such as with a gradient fill type.

## Property Details

### backgroundColor

Specifies the shape fill background color. You can provide the value in the '#RRGGBB' format or the color name.

```typescript
backgroundColor?: string;
```

- Property Value: string  
- Remarks: [ API set: [WordApiDesktop 1.2](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

### foregroundColor

Specifies the shape fill foreground color. You can provide the value in the '#RRGGBB' format or the color name.

```typescript
foregroundColor?: string;
```

- Property Value: string  
- Remarks: [ API set: [WordApiDesktop 1.2](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

### transparency

Specifies the transparency percentage of the fill as a value from 0.0 (opaque) through 1.0 (clear). Returns `null` if the shape type does not support transparency or the shape fill has inconsistent transparency, such as with a gradient fill type.

```typescript
transparency?: number;
```

- Property Value: number  
- Remarks: [ API set: [WordApiDesktop 1.2](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]