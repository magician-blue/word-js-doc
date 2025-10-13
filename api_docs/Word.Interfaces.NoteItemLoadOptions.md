# Word.Interfaces.NoteItemLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Represents a footnote or endnote.

## Remarks

[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- $all  
  Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

- body  
  Represents the body object of the note item. It's the portion of the text within the footnote or endnote.

- reference  
  Represents a footnote or endnote reference in the main document.

- type  
  Represents the note item type: footnote or endnote.

## Property Details

### $all

Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

Property Value: boolean

---

### body

Represents the body object of the note item. It's the portion of the text within the footnote or endnote.

```typescript
body?: Word.Interfaces.BodyLoadOptions;
```

Property Value: [Word.Interfaces.BodyLoadOptions](/en-us/javascript/api/word/word.interfaces.bodyloadoptions)

Remarks: [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### reference

Represents a footnote or endnote reference in the main document.

```typescript
reference?: Word.Interfaces.RangeLoadOptions;
```

Property Value: [Word.Interfaces.RangeLoadOptions](/en-us/javascript/api/word/word.interfaces.rangeloadoptions)

Remarks: [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### type

Represents the note item type: footnote or endnote.

```typescript
type?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)