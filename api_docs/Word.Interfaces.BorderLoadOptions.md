# Word.Interfaces.BorderLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Represents the Border object for text, a paragraph, or a table.

## Remarks

[API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- $all  
  Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).

- color  
  Specifies the color for the border. Color is specified in '#RRGGBB' format or by using the color name.

- location  
  Gets the location of the border.

- type  
  Specifies the border type for the border.

- visible  
  Specifies whether the border is visible.

- width  
  Specifies the width for the border.

## Property Details

### $all

Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).

```typescript
$all?: boolean;
```

Property Value  
boolean

### color

Specifies the color for the border. Color is specified in '#RRGGBB' format or by using the color name.

```typescript
color?: boolean;
```

Property Value  
boolean

Remarks  
[API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### location

Gets the location of the border.

```typescript
location?: boolean;
```

Property Value  
boolean

Remarks  
[API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### type

Specifies the border type for the border.

```typescript
type?: boolean;
```

Property Value  
boolean

Remarks  
[API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### visible

Specifies whether the border is visible.

```typescript
visible?: boolean;
```

Property Value  
boolean

Remarks  
[API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### width

Specifies the width for the border.

```typescript
width?: boolean;
```

Property Value  
boolean

Remarks  
[API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)