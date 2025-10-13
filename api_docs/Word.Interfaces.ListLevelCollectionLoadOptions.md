# Word.Interfaces.ListLevelCollectionLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Contains a collection of [Word.ListLevel](/en-us/javascript/api/word/word.listlevel) objects.

## Remarks

[ API set: WordApiDesktop 1.1 ]

## Properties

- $all: Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).
- alignment: For EACH ITEM in the collection: Specifies the horizontal alignment of the list level. The value can be 'Left', 'Centered', or 'Right'.
- font: For EACH ITEM in the collection: Gets a Font object that represents the character formatting of the specified object.
- linkedStyle: For EACH ITEM in the collection: Specifies the name of the style that's linked to the specified list level object.
- numberFormat: For EACH ITEM in the collection: Specifies the number format for the specified list level.
- numberPosition: For EACH ITEM in the collection: Specifies the position (in points) of the number or bullet for the specified list level object.
- numberStyle: For EACH ITEM in the collection: Specifies the number style for the list level object.
- resetOnHigher: For EACH ITEM in the collection: Specifies the list level that must appear before the specified list level restarts numbering at 1.
- startAt: For EACH ITEM in the collection: Specifies the starting number for the specified list level object.
- tabPosition: For EACH ITEM in the collection: Specifies the tab position for the specified list level object.
- textPosition: For EACH ITEM in the collection: Specifies the position (in points) for the second line of wrapping text for the specified list level object.
- trailingCharacter: For EACH ITEM in the collection: Specifies the character inserted after the number for the specified list level.

## Property Details

### $all

Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).

```typescript
$all?: boolean;
```

Property Value: boolean

---

### alignment

For EACH ITEM in the collection: Specifies the horizontal alignment of the list level. The value can be 'Left', 'Centered', or 'Right'.

```typescript
alignment?: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApiDesktop 1.1 ]

---

### font

For EACH ITEM in the collection: Gets a Font object that represents the character formatting of the specified object.

```typescript
font?: Word.Interfaces.FontLoadOptions;
```

Property Value: [Word.Interfaces.FontLoadOptions](/en-us/javascript/api/word/word.interfaces.fontloadoptions)

Remarks: [ API set: WordApiDesktop 1.1 ]

---

### linkedStyle

For EACH ITEM in the collection: Specifies the name of the style that's linked to the specified list level object.

```typescript
linkedStyle?: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApiDesktop 1.1 ]

---

### numberFormat

For EACH ITEM in the collection: Specifies the number format for the specified list level.

```typescript
numberFormat?: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApiDesktop 1.1 ]

---

### numberPosition

For EACH ITEM in the collection: Specifies the position (in points) of the number or bullet for the specified list level object.

```typescript
numberPosition?: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApiDesktop 1.1 ]

---

### numberStyle

For EACH ITEM in the collection: Specifies the number style for the list level object.

```typescript
numberStyle?: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApiDesktop 1.1 ]

---

### resetOnHigher

For EACH ITEM in the collection: Specifies the list level that must appear before the specified list level restarts numbering at 1.

```typescript
resetOnHigher?: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApiDesktop 1.1 ]

---

### startAt

For EACH ITEM in the collection: Specifies the starting number for the specified list level object.

```typescript
startAt?: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApiDesktop 1.1 ]

---

### tabPosition

For EACH ITEM in the collection: Specifies the tab position for the specified list level object.

```typescript
tabPosition?: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApiDesktop 1.1 ]

---

### textPosition

For EACH ITEM in the collection: Specifies the position (in points) for the second line of wrapping text for the specified list level object.

```typescript
textPosition?: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApiDesktop 1.1 ]

---

### trailingCharacter

For EACH ITEM in the collection: Specifies the character inserted after the number for the specified list level.

```typescript
trailingCharacter?: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApiDesktop 1.1 ]