# Word.Interfaces.BodyLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Represents the body of a document or a section.

## Remarks

[API set: WordApi 1.1]

## Properties

- `$all`: Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
- `font`: Gets the text format of the body. Use this to get and set font name, size, color and other properties.
- `parentBody`: Gets the parent body of the body. For example, a table cell body's parent body could be a header. Throws an `ItemNotFound` error if there isn't a parent body.
- `parentBodyOrNullObject`: Gets the parent body of the body. For example, a table cell body's parent body could be a header. If there isn't a parent body, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).
- `parentContentControl`: Gets the content control that contains the body. Throws an `ItemNotFound` error if there isn't a parent content control.
- `parentContentControlOrNullObject`: Gets the content control that contains the body. If there isn't a parent content control, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).
- `parentSection`: Gets the parent section of the body. Throws an `ItemNotFound` error if there isn't a parent section.
- `parentSectionOrNullObject`: Gets the parent section of the body. If there isn't a parent section, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).
- `style`: Specifies the style name for the body. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
- `styleBuiltIn`: Specifies the built-in style name for the body. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
- `text`: Gets the text of the body. Use the insertText method to insert text.
- `type`: Gets the type of the body. The type can be 'MainDoc', 'Section', 'Header', 'Footer', or 'TableCell'. Additional types âFootnoteâ, âEndnoteâ, and âNoteItemâ are supported in WordAPIOnline 1.1 and later.

## Property Details

### $all

Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

- Type: boolean

---

### font

Gets the text format of the body. Use this to get and set font name, size, color and other properties.

```typescript
font?: Word.Interfaces.FontLoadOptions;
```

- Type: [Word.Interfaces.FontLoadOptions](/en-us/javascript/api/word/word.interfaces.fontloadoptions)
- Remarks: [API set: WordApi 1.1]

---

### parentBody

Gets the parent body of the body. For example, a table cell body's parent body could be a header. Throws an `ItemNotFound` error if there isn't a parent body.

```typescript
parentBody?: Word.Interfaces.BodyLoadOptions;
```

- Type: [Word.Interfaces.BodyLoadOptions](/en-us/javascript/api/word/word.interfaces.bodyloadoptions)
- Remarks: [API set: WordApi 1.3]

---

### parentBodyOrNullObject

Gets the parent body of the body. For example, a table cell body's parent body could be a header. If there isn't a parent body, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
parentBodyOrNullObject?: Word.Interfaces.BodyLoadOptions;
```

- Type: [Word.Interfaces.BodyLoadOptions](/en-us/javascript/api/word/word.interfaces.bodyloadoptions)
- Remarks: [API set: WordApi 1.3]

---

### parentContentControl

Gets the content control that contains the body. Throws an `ItemNotFound` error if there isn't a parent content control.

```typescript
parentContentControl?: Word.Interfaces.ContentControlLoadOptions;
```

- Type: [Word.Interfaces.ContentControlLoadOptions](/en-us/javascript/api/word/word.interfaces.contentcontrolloadoptions)
- Remarks: [API set: WordApi 1.1]

---

### parentContentControlOrNullObject

Gets the content control that contains the body. If there isn't a parent content control, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
parentContentControlOrNullObject?: Word.Interfaces.ContentControlLoadOptions;
```

- Type: [Word.Interfaces.ContentControlLoadOptions](/en-us/javascript/api/word/word.interfaces.contentcontrolloadoptions)
- Remarks: [API set: WordApi 1.3]

---

### parentSection

Gets the parent section of the body. Throws an `ItemNotFound` error if there isn't a parent section.

```typescript
parentSection?: Word.Interfaces.SectionLoadOptions;
```

- Type: [Word.Interfaces.SectionLoadOptions](/en-us/javascript/api/word/word.interfaces.sectionloadoptions)
- Remarks: [API set: WordApi 1.3]

---

### parentSectionOrNullObject

Gets the parent section of the body. If there isn't a parent section, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
parentSectionOrNullObject?: Word.Interfaces.SectionLoadOptions;
```

- Type: [Word.Interfaces.SectionLoadOptions](/en-us/javascript/api/word/word.interfaces.sectionloadoptions)
- Remarks: [API set: WordApi 1.3]

---

### style

Specifies the style name for the body. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.

```typescript
style?: boolean;
```

- Type: boolean
- Remarks: [API set: WordApi 1.1]

---

### styleBuiltIn

Specifies the built-in style name for the body. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.

```typescript
styleBuiltIn?: boolean;
```

- Type: boolean
- Remarks: [API set: WordApi 1.3]

---

### text

Gets the text of the body. Use the insertText method to insert text.

```typescript
text?: boolean;
```

- Type: boolean
- Remarks: [API set: WordApi 1.1]

---

### type

Gets the type of the body. The type can be 'MainDoc', 'Section', 'Header', 'Footer', or 'TableCell'. Additional types âFootnoteâ, âEndnoteâ, and âNoteItemâ are supported in WordAPIOnline 1.1 and later.

```typescript
type?: boolean;
```

- Type: boolean
- Remarks: [API set: WordApi 1.3]