# Word.Interfaces.TableBorderLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Specifies the border style.

## Remarks

[API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- $all: Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
- color: Specifies the table border color.
- type: Specifies the type of the table border.
- width: Specifies the width, in points, of the table border. Not applicable to table border types that have fixed widths.

## Property Details

### $all

Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

Property Value: boolean

### color

Specifies the table border color.

```typescript
color?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### type

Specifies the type of the table border.

```typescript
type?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### width

Specifies the width, in points, of the table border. Not applicable to table border types that have fixed widths.

```typescript
width?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)