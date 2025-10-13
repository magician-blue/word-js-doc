# Word.Interfaces.ListLevelLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Represents a list level.

## Remarks

[API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- $all — Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).
- alignment — Specifies the horizontal alignment of the list level. The value can be 'Left', 'Centered', or 'Right'.
- font — Gets a Font object that represents the character formatting of the specified object.
- linkedStyle — Specifies the name of the style that's linked to the specified list level object.
- numberFormat — Specifies the number format for the specified list level.
- numberPosition — Specifies the position (in points) of the number or bullet for the specified list level object.
- numberStyle — Specifies the number style for the list level object.
- resetOnHigher — Specifies the list level that must appear before the specified list level restarts numbering at 1.
- startAt — Specifies the starting number for the specified list level object.
- tabPosition — Specifies the tab position for the specified list level object.
- textPosition — Specifies the position (in points) for the second line of wrapping text for the specified list level object.
- trailingCharacter — Specifies the character inserted after the number for the specified list level.

## Property Details

### $all

Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).

```typescript
$all?: boolean;
```

Property Value
- boolean

### alignment

Specifies the horizontal alignment of the list level. The value can be 'Left', 'Centered', or 'Right'.

```typescript
alignment?: boolean;
```

Property Value
- boolean

Remarks  
[API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### font

Gets a Font object that represents the character formatting of the specified object.

```typescript
font?: Word.Interfaces.FontLoadOptions;
```

Property Value
- [Word.Interfaces.FontLoadOptions](/en-us/javascript/api/word/word.interfaces.fontloadoptions)

Remarks  
[API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### linkedStyle

Specifies the name of the style that's linked to the specified list level object.

```typescript
linkedStyle?: boolean;
```

Property Value
- boolean

Remarks  
[API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### numberFormat

Specifies the number format for the specified list level.

```typescript
numberFormat?: boolean;
```

Property Value
- boolean

Remarks  
[API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### numberPosition

Specifies the position (in points) of the number or bullet for the specified list level object.

```typescript
numberPosition?: boolean;
```

Property Value
- boolean

Remarks  
[API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### numberStyle

Specifies the number style for the list level object.

```typescript
numberStyle?: boolean;
```

Property Value
- boolean

Remarks  
[API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### resetOnHigher

Specifies the list level that must appear before the specified list level restarts numbering at 1.

```typescript
resetOnHigher?: boolean;
```

Property Value
- boolean

Remarks  
[API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### startAt

Specifies the starting number for the specified list level object.

```typescript
startAt?: boolean;
```

Property Value
- boolean

Remarks  
[API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### tabPosition

Specifies the tab position for the specified list level object.

```typescript
tabPosition?: boolean;
```

Property Value
- boolean

Remarks  
[API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### textPosition

Specifies the position (in points) for the second line of wrapping text for the specified list level object.

```typescript
textPosition?: boolean;
```

Property Value
- boolean

Remarks  
[API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### trailingCharacter

Specifies the character inserted after the number for the specified list level.

```typescript
trailingCharacter?: boolean;
```

Property Value
- boolean

Remarks  
[API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)