# Word.Interfaces.CanvasLoadOptions interface

- Package: https://learn.microsoft.com/en-us/javascript/api/word

Represents a canvas in the document. To get the corresponding Shape object, use Canvas.shape.

## Remarks

[ API set: WordApiDesktop 1.2 ](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- $all: Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).
- id: Gets an integer that represents the canvas identifier.
- shape: Gets the Shape object associated with the canvas.

## Property Details

### $all

Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).

```typescript
$all?: boolean;
```

Property Value: boolean

### id

Gets an integer that represents the canvas identifier.

```typescript
id?: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApiDesktop 1.2 ](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### shape

Gets the Shape object associated with the canvas.

```typescript
shape?: Word.Interfaces.ShapeLoadOptions;
```

Property Value: https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.shapeloadoptions

Remarks: [ API set: WordApiDesktop 1.2 ](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)