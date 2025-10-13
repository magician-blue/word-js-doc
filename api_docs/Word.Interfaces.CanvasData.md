# Word.Interfaces.CanvasData interface

Package: [word](/en-us/javascript/api/word)

An interface describing the data returned by calling `canvas.toJSON()`.

## Properties

- id: Gets an integer that represents the canvas identifier.
- shape: Gets the Shape object associated with the canvas.
- shapes: Gets the collection of Shape objects. Currently, only text boxes, pictures, and geometric shapes are supported.

## Property Details

### id

Gets an integer that represents the canvas identifier.

```typescript
id?: number;
```

Property Value
- number

Remarks
- [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### shape

Gets the Shape object associated with the canvas.

```typescript
shape?: Word.Interfaces.ShapeData;
```

Property Value
- [Word.Interfaces.ShapeData](/en-us/javascript/api/word/word.interfaces.shapedata)

Remarks
- [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### shapes

Gets the collection of Shape objects. Currently, only text boxes, pictures, and geometric shapes are supported.

```typescript
shapes?: Word.Interfaces.ShapeData[];
```

Property Value
- [Word.Interfaces.ShapeData](/en-us/javascript/api/word/word.interfaces.shapedata)[]

Remarks
- [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)