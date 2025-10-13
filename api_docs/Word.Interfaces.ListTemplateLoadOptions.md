# Word.Interfaces.ListTemplateLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Represents a list template.

## Remarks

[ API set: WordApiDesktop 1.1 ]

## Properties

- $all  
  Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

- outlineNumbered  
  Specifies whether the list template is outline numbered.

## Property Details

### $all

Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

Property Value  
boolean

### outlineNumbered

Specifies whether the list template is outline numbered.

```typescript
outlineNumbered?: boolean;
```

Property Value  
boolean

Remarks  
[ API set: WordApiDesktop 1.1 ]