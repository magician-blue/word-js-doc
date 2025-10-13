# Word.Interfaces.ShapeGroupData interface

Package: https://learn.microsoft.com/en-us/javascript/api/word

An interface describing the data returned by calling `shapeGroup.toJSON()`.

## Properties

- [id](#id) — Gets an integer that represents the shape group identifier.
- [shape](#shape) — Gets the Shape object associated with the group.
- [shapes](#shapes) — Gets the collection of Shape objects. Currently, only text boxes, geometric shapes, and pictures are supported.

## Property Details

### id

Gets an integer that represents the shape group identifier.

`id?: number;`

Property value: number

Remarks: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets (API set: WordApiDesktop 1.2)

### shape

Gets the Shape object associated with the group.

`shape?: Word.Interfaces.ShapeData;`

Property value: https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.shapedata (Word.Interfaces.ShapeData)

Remarks: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets (API set: WordApiDesktop 1.2)

### shapes

Gets the collection of Shape objects. Currently, only text boxes, geometric shapes, and pictures are supported.

`shapes?: Word.Interfaces.ShapeData[];`

Property value: https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.shapedata (Word.Interfaces.ShapeData)[]

Remarks: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets (API set: WordApiDesktop 1.2)