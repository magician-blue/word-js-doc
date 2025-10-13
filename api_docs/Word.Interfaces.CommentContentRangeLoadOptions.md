# Word.Interfaces.CommentContentRangeLoadOptions interface

Package: [word](https://learn.microsoft.com/en-us/javascript/api/word)

## Remarks

[ API set: WordApi 1.4 ]

## Properties

- $all
  - Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
- bold
  - Specifies a value that indicates whether the comment text is bold.
- hyperlink
  - Gets the first hyperlink in the range, or sets a hyperlink on the range. All hyperlinks in the range are deleted when you set a new hyperlink on the range.
- isEmpty
  - Checks whether the range length is zero.
- italic
  - Specifies a value that indicates whether the comment text is italicized.
- strikeThrough
  - Specifies a value that indicates whether the comment text has a strikethrough.
- text
  - Gets the text of the comment range.
- underline
  - Specifies a value that indicates the comment text's underline type. 'None' if the comment text isn't underlined.

## Property Details

### $all

Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

Property Value: boolean

### bold

Specifies a value that indicates whether the comment text is bold.

```typescript
bold?: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApi 1.4 ]

### hyperlink

Gets the first hyperlink in the range, or sets a hyperlink on the range. All hyperlinks in the range are deleted when you set a new hyperlink on the range.

```typescript
hyperlink?: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApi 1.4 ]

### isEmpty

Checks whether the range length is zero.

```typescript
isEmpty?: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApi 1.4 ]

### italic

Specifies a value that indicates whether the comment text is italicized.

```typescript
italic?: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApi 1.4 ]

### strikeThrough

Specifies a value that indicates whether the comment text has a strikethrough.

```typescript
strikeThrough?: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApi 1.4 ]

### text

Gets the text of the comment range.

```typescript
text?: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApi 1.4 ]

### underline

Specifies a value that indicates the comment text's underline type. 'None' if the comment text isn't underlined.

```typescript
underline?: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApi 1.4 ]