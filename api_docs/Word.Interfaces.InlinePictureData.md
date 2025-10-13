# Word.Interfaces.InlinePictureData interface

Package: [word](/en-us/javascript/api/word)

An interface describing the data returned by calling `inlinePicture.toJSON()`.

## Properties

- `altTextDescription` — Specifies a string that represents the alternative text associated with the inline image.
- `altTextTitle` — Specifies a string that contains the title for the inline image.
- `height` — Specifies a number that describes the height of the inline image.
- `hyperlink` — Specifies a hyperlink on the image. Use a '#' to separate the address part from the optional location part.
- `imageFormat` — Gets the format of the inline image.
- `lockAspectRatio` — Specifies a value that indicates whether the inline image retains its original proportions when you resize it.
- `width` — Specifies a number that describes the width of the inline image.

## Property Details

### altTextDescription

Specifies a string that represents the alternative text associated with the inline image.

```typescript
altTextDescription?: string;
```

#### Property Value
- string

#### Remarks
[API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### altTextTitle

Specifies a string that contains the title for the inline image.

```typescript
altTextTitle?: string;
```

#### Property Value
- string

#### Remarks
[API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### height

Specifies a number that describes the height of the inline image.

```typescript
height?: number;
```

#### Property Value
- number

#### Remarks
[API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### hyperlink

Specifies a hyperlink on the image. Use a '#' to separate the address part from the optional location part.

```typescript
hyperlink?: string;
```

#### Property Value
- string

#### Remarks
[API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### imageFormat

Gets the format of the inline image.

```typescript
imageFormat?: Word.ImageFormat | "Unsupported" | "Undefined" | "Bmp" | "Jpeg" | "Gif" | "Tiff" | "Png" | "Icon" | "Exif" | "Wmf" | "Emf" | "Pict" | "Pdf" | "Svg";
```

#### Property Value
- [Word.ImageFormat](/en-us/javascript/api/word/word.imageformat) | "Unsupported" | "Undefined" | "Bmp" | "Jpeg" | "Gif" | "Tiff" | "Png" | "Icon" | "Exif" | "Wmf" | "Emf" | "Pict" | "Pdf" | "Svg"

#### Remarks
[API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### lockAspectRatio

Specifies a value that indicates whether the inline image retains its original proportions when you resize it.

```typescript
lockAspectRatio?: boolean;
```

#### Property Value
- boolean

#### Remarks
[API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### width

Specifies a number that describes the width of the inline image.

```typescript
width?: number;
```

#### Property Value
- number

#### Remarks
[API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)