# Word.Interfaces.BorderCollectionLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Represents the collection of border styles.

## Remarks

[ API set: WordApiDesktop 1.1 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- $all  
  Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

- color  
  For EACH ITEM in the collection: Specifies the color for the border. Color is specified in â#RRGGBBâ format or by using the color name.

- location  
  For EACH ITEM in the collection: Gets the location of the border.

- type  
  For EACH ITEM in the collection: Specifies the border type for the border.

- visible  
  For EACH ITEM in the collection: Specifies whether the border is visible.

- width  
  For EACH ITEM in the collection: Specifies the width for the border.

## Property Details

### $all

Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

Property Value: boolean

### color

For EACH ITEM in the collection: Specifies the color for the border. Color is specified in â#RRGGBBâ format or by using the color name.

```typescript
color?: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApiDesktop 1.1 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### location

For EACH ITEM in the collection: Gets the location of the border.

```typescript
location?: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApiDesktop 1.1 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### type

For EACH ITEM in the collection: Specifies the border type for the border.

```typescript
type?: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApiDesktop 1.1 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### visible

For EACH ITEM in the collection: Specifies whether the border is visible.

```typescript
visible?: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApiDesktop 1.1 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### width

For EACH ITEM in the collection: Specifies the width for the border.

```typescript
width?: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApiDesktop 1.1 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)