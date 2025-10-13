# Word.Interfaces.ListLevelData interface

Package: [word](https://learn.microsoft.com/en-us/javascript/api/word)

An interface describing the data returned by calling `listLevel.toJSON()`.

## Properties

- [alignment](#alignment) — Specifies the horizontal alignment of the list level. The value can be 'Left', 'Centered', or 'Right'.
- [font](#font) — Gets a Font object that represents the character formatting of the specified object.
- [linkedStyle](#linkedstyle) — Specifies the name of the style that's linked to the specified list level object.
- [numberFormat](#numberformat) — Specifies the number format for the specified list level.
- [numberPosition](#numberposition) — Specifies the position (in points) of the number or bullet for the specified list level object.
- [numberStyle](#numberstyle) — Specifies the number style for the list level object.
- [resetOnHigher](#resetonhigher) — Specifies the list level that must appear before the specified list level restarts numbering at 1.
- [startAt](#startat) — Specifies the starting number for the specified list level object.
- [tabPosition](#tabposition) — Specifies the tab position for the specified list level object.
- [textPosition](#textposition) — Specifies the position (in points) for the second line of wrapping text for the specified list level object.
- [trailingCharacter](#trailingcharacter) — Specifies the character inserted after the number for the specified list level.

## Property Details

### alignment

Specifies the horizontal alignment of the list level. The value can be 'Left', 'Centered', or 'Right'.

```typescript
alignment?: Word.Alignment | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified";
```

- Property Value: [Word.Alignment](https://learn.microsoft.com/en-us/javascript/api/word/word.alignment) | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified"

Remarks: [API set: WordApiDesktop 1.1](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### font

Gets a Font object that represents the character formatting of the specified object.

```typescript
font?: Word.Interfaces.FontData;
```

- Property Value: [Word.Interfaces.FontData](https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.fontdata)

Remarks: [API set: WordApiDesktop 1.1](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### linkedStyle

Specifies the name of the style that's linked to the specified list level object.

```typescript
linkedStyle?: string;
```

- Property Value: string

Remarks: [API set: WordApiDesktop 1.1](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### numberFormat

Specifies the number format for the specified list level.

```typescript
numberFormat?: string;
```

- Property Value: string

Remarks: [API set: WordApiDesktop 1.1](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### numberPosition

Specifies the position (in points) of the number or bullet for the specified list level object.

```typescript
numberPosition?: number;
```

- Property Value: number

Remarks: [API set: WordApiDesktop 1.1](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### numberStyle

Specifies the number style for the list level object.

```typescript
numberStyle?: Word.ListBuiltInNumberStyle | "None" | "Arabic" | "UpperRoman" | "LowerRoman" | "UpperLetter" | "LowerLetter" | "Ordinal" | "CardinalText" | "OrdinalText" | "Kanji" | "KanjiDigit" | "AiueoHalfWidth" | "IrohaHalfWidth" | "ArabicFullWidth" | "KanjiTraditional" | "KanjiTraditional2" | "NumberInCircle" | "Aiueo" | "Iroha" | "ArabicLZ" | "Bullet" | "Ganada" | "Chosung" | "GBNum1" | "GBNum2" | "GBNum3" | "GBNum4" | "Zodiac1" | "Zodiac2" | "Zodiac3" | "TradChinNum1" | "TradChinNum2" | "TradChinNum3" | "TradChinNum4" | "SimpChinNum1" | "SimpChinNum2" | "SimpChinNum3" | "SimpChinNum4" | "HanjaRead" | "HanjaReadDigit" | "Hangul" | "Hanja" | "Hebrew1" | "Arabic1" | "Hebrew2" | "Arabic2" | "HindiLetter1" | "HindiLetter2" | "HindiArabic" | "HindiCardinalText" | "ThaiLetter" | "ThaiArabic" | "ThaiCardinalText" | "VietCardinalText" | "LowercaseRussian" | "UppercaseRussian" | "LowercaseGreek" | "UppercaseGreek" | "ArabicLZ2" | "ArabicLZ3" | "ArabicLZ4" | "LowercaseTurkish" | "UppercaseTurkish" | "LowercaseBulgarian" | "UppercaseBulgarian" | "PictureBullet" | "Legal" | "LegalLZ";
```

- Property Value: [Word.ListBuiltInNumberStyle](https://learn.microsoft.com/en-us/javascript/api/word/word.listbuiltinnumberstyle) | "None" | "Arabic" | "UpperRoman" | "LowerRoman" | "UpperLetter" | "LowerLetter" | "Ordinal" | "CardinalText" | "OrdinalText" | "Kanji" | "KanjiDigit" | "AiueoHalfWidth" | "IrohaHalfWidth" | "ArabicFullWidth" | "KanjiTraditional" | "KanjiTraditional2" | "NumberInCircle" | "Aiueo" | "Iroha" | "ArabicLZ" | "Bullet" | "Ganada" | "Chosung" | "GBNum1" | "GBNum2" | "GBNum3" | "GBNum4" | "Zodiac1" | "Zodiac2" | "Zodiac3" | "TradChinNum1" | "TradChinNum2" | "TradChinNum3" | "TradChinNum4" | "SimpChinNum1" | "SimpChinNum2" | "SimpChinNum3" | "SimpChinNum4" | "HanjaRead" | "HanjaReadDigit" | "Hangul" | "Hanja" | "Hebrew1" | "Arabic1" | "Hebrew2" | "Arabic2" | "HindiLetter1" | "HindiLetter2" | "HindiArabic" | "HindiCardinalText" | "ThaiLetter" | "ThaiArabic" | "ThaiCardinalText" | "VietCardinalText" | "LowercaseRussian" | "UppercaseRussian" | "LowercaseGreek" | "UppercaseGreek" | "ArabicLZ2" | "ArabicLZ3" | "ArabicLZ4" | "LowercaseTurkish" | "UppercaseTurkish" | "LowercaseBulgarian" | "UppercaseBulgarian" | "PictureBullet" | "Legal" | "LegalLZ"

Remarks: [API set: WordApiDesktop 1.1](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### resetOnHigher

Specifies the list level that must appear before the specified list level restarts numbering at 1.

```typescript
resetOnHigher?: number;
```

- Property Value: number

Remarks: [API set: WordApiDesktop 1.1](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### startAt

Specifies the starting number for the specified list level object.

```typescript
startAt?: number;
```

- Property Value: number

Remarks: [API set: WordApiDesktop 1.1](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### tabPosition

Specifies the tab position for the specified list level object.

```typescript
tabPosition?: number;
```

- Property Value: number

Remarks: [API set: WordApiDesktop 1.1](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### textPosition

Specifies the position (in points) for the second line of wrapping text for the specified list level object.

```typescript
textPosition?: number;
```

- Property Value: number

Remarks: [API set: WordApiDesktop 1.1](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### trailingCharacter

Specifies the character inserted after the number for the specified list level.

```typescript
trailingCharacter?: Word.TrailingCharacter | "TrailingTab" | "TrailingSpace" | "TrailingNone";
```

- Property Value: [Word.TrailingCharacter](https://learn.microsoft.com/en-us/javascript/api/word/word.trailingcharacter) | "TrailingTab" | "TrailingSpace" | "TrailingNone"

Remarks: [API set: WordApiDesktop 1.1](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)