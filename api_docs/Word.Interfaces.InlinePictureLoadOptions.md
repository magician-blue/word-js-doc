# Word.Interfaces.InlinePictureLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Represents an inline picture.

## Remarks

API set: [WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- $all
  - Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

- altTextDescription
  - Specifies a string that represents the alternative text associated with the inline image.

- altTextTitle
  - Specifies a string that contains the title for the inline image.

- height
  - Specifies a number that describes the height of the inline image.

- hyperlink
  - Specifies a hyperlink on the image. Use a '#' to separate the address part from the optional location part.

- imageFormat
  - Gets the format of the inline image.

- lockAspectRatio
  - Specifies a value that indicates whether the inline image retains its original proportions when you resize it.

- paragraph
  - Gets the parent paragraph that contains the inline image.

- parentContentControl
  - Gets the content control that contains the inline image. Throws an `ItemNotFound` error if there isn't a parent content control.

- parentContentControlOrNullObject
  - Gets the content control that contains the inline image. If there isn't a parent content control, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [\*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

- parentTable
  - Gets the table that contains the inline image. Throws an `ItemNotFound` error if it isn't contained in a table.

- parentTableCell
  - Gets the table cell that contains the inline image. Throws an `ItemNotFound` error if it isn't contained in a table cell.

- parentTableCellOrNullObject
  - Gets the table cell that contains the inline image. If it isn't contained in a table cell, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [\*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

- parentTableOrNullObject
  - Gets the table that contains the inline image. If it isn't contained in a table, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [\*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

- width
  - Specifies a number that describes the width of the inline image.

## Property Details

### $all

Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

Property Value
- boolean

---

### altTextDescription

Specifies a string that represents the alternative text associated with the inline image.

```typescript
altTextDescription?: boolean;
```

Property Value
- boolean

Remarks
- API set: [WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### altTextTitle

Specifies a string that contains the title for the inline image.

```typescript
altTextTitle?: boolean;
```

Property Value
- boolean

Remarks
- API set: [WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### height

Specifies a number that describes the height of the inline image.

```typescript
height?: boolean;
```

Property Value
- boolean

Remarks
- API set: [WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### hyperlink

Specifies a hyperlink on the image. Use a '#' to separate the address part from the optional location part.

```typescript
hyperlink?: boolean;
```

Property Value
- boolean

Remarks
- API set: [WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### imageFormat

Gets the format of the inline image.

```typescript
imageFormat?: boolean;
```

Property Value
- boolean

Remarks
- API set: [WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### lockAspectRatio

Specifies a value that indicates whether the inline image retains its original proportions when you resize it.

```typescript
lockAspectRatio?: boolean;
```

Property Value
- boolean

Remarks
- API set: [WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### paragraph

Gets the parent paragraph that contains the inline image.

```typescript
paragraph?: Word.Interfaces.ParagraphLoadOptions;
```

Property Value
- [Word.Interfaces.ParagraphLoadOptions](/en-us/javascript/api/word/word.interfaces.paragraphloadoptions)

Remarks
- API set: [WordApi 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### parentContentControl

Gets the content control that contains the inline image. Throws an `ItemNotFound` error if there isn't a parent content control.

```typescript
parentContentControl?: Word.Interfaces.ContentControlLoadOptions;
```

Property Value
- [Word.Interfaces.ContentControlLoadOptions](/en-us/javascript/api/word/word.interfaces.contentcontrolloadoptions)

Remarks
- API set: [WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### parentContentControlOrNullObject

Gets the content control that contains the inline image. If there isn't a parent content control, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [\*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
parentContentControlOrNullObject?: Word.Interfaces.ContentControlLoadOptions;
```

Property Value
- [Word.Interfaces.ContentControlLoadOptions](/en-us/javascript/api/word/word.interfaces.contentcontrolloadoptions)

Remarks
- API set: [WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### parentTable

Gets the table that contains the inline image. Throws an `ItemNotFound` error if it isn't contained in a table.

```typescript
parentTable?: Word.Interfaces.TableLoadOptions;
```

Property Value
- [Word.Interfaces.TableLoadOptions](/en-us/javascript/api/word/word.interfaces.tableloadoptions)

Remarks
- API set: [WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### parentTableCell

Gets the table cell that contains the inline image. Throws an `ItemNotFound` error if it isn't contained in a table cell.

```typescript
parentTableCell?: Word.Interfaces.TableCellLoadOptions;
```

Property Value
- [Word.Interfaces.TableCellLoadOptions](/en-us/javascript/api/word/word.interfaces.tablecellloadoptions)

Remarks
- API set: [WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### parentTableCellOrNullObject

Gets the table cell that contains the inline image. If it isn't contained in a table cell, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [\*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
parentTableCellOrNullObject?: Word.Interfaces.TableCellLoadOptions;
```

Property Value
- [Word.Interfaces.TableCellLoadOptions](/en-us/javascript/api/word/word.interfaces.tablecellloadoptions)

Remarks
- API set: [WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### parentTableOrNullObject

Gets the table that contains the inline image. If it isn't contained in a table, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [\*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
parentTableOrNullObject?: Word.Interfaces.TableLoadOptions;
```

Property Value
- [Word.Interfaces.TableLoadOptions](/en-us/javascript/api/word/word.interfaces.tableloadoptions)

Remarks
- API set: [WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### width

Specifies a number that describes the width of the inline image.

```typescript
width?: boolean;
```

Property Value
- boolean

Remarks
- API set: [WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)