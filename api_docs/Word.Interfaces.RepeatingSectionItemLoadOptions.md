# Word.Interfaces.RepeatingSectionItemLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents a single item in a Word.RepeatingSectionContentControl.

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

## Properties

- $all  
  Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).

- range  
  Returns the range of this repeating section item, excluding the start and end tags.

## Property Details

### $all

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).

```typescript
$all?: boolean;
```

Property Value
boolean

### range

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the range of this repeating section item, excluding the start and end tags.

```typescript
range?: Word.Interfaces.RangeLoadOptions;
```

Property Value
[Word.Interfaces.RangeLoadOptions](/en-us/javascript/api/word/word.interfaces.rangeloadoptions)

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]