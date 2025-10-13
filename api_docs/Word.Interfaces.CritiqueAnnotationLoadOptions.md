# Word.Interfaces.CritiqueAnnotationLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Represents an annotation wrapper around critique displayed in the document.

## Remarks

[API set: WordApi 1.7](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- $all  
  Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

- critique  
  Gets the critique that was passed when the annotation was inserted.

- range  
  Gets the range of text that is annotated.

## Property Details

### $all

Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

Property Value
- boolean

### critique

Gets the critique that was passed when the annotation was inserted.

```typescript
critique?: boolean;
```

Property Value
- boolean

Remarks  
[API set: WordApi 1.7](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### range

Gets the range of text that is annotated.

```typescript
range?: Word.Interfaces.RangeLoadOptions;
```

Property Value
- [Word.Interfaces.RangeLoadOptions](/en-us/javascript/api/word/word.interfaces.rangeloadoptions)

Remarks  
[API set: WordApi 1.7](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)