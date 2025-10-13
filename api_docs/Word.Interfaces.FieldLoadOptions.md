# Word.Interfaces.FieldLoadOptions interface

Package: word

Summary
- Represents a field.

## Remarks
- [API set: WordApi 1.4]
- Important: To learn more about which fields can be inserted, see the Word.Range.insertField API introduced in requirement set 1.5. Support for managing fields is similar to what's available in the Word UI. However, the Word UI on the web primarily only supports fields as read-only (see Field codes in Word for the web: https://support.microsoft.com/office/d8f46094-13c3-4966-98c3-259748f3caf1). To learn more about Word UI clients that more fully support fields, see the product list at the beginning of Insert, edit, and view fields in Word: https://support.microsoft.com/office/c429bbb0-8669-48a7-bd24-bab6ba6b06bb.

## Properties
- $all — Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).
- code — Specifies the field's code instruction.
- data — Specifies data in an "Addin" field. If the field isn't an "Addin" field, it is null and it will throw a general exception when code attempts to set it.
- kind — Gets the field's kind.
- linkFormat — Gets a LinkFormat object that represents the link options of the field.
- locked — Specifies whether the field is locked. true if the field is locked, false otherwise.
- oleFormat — Gets an OleFormat object that represents the OLE characteristics (other than linking) for the field.
- parentBody — Gets the parent body of the field.
- parentContentControl — Gets the content control that contains the field. Throws an ItemNotFound error if there isn't a parent content control.
- parentContentControlOrNullObject — Gets the content control that contains the field. If there isn't a parent content control, then this method will return an object with its isNullObject property set to true. For further information, see “OrNullObject methods and properties”: https://learn.microsoft.com/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties.
- parentTable — Gets the table that contains the field. Throws an ItemNotFound error if it isn't contained in a table.
- parentTableCell — Gets the table cell that contains the field. Throws an ItemNotFound error if it isn't contained in a table cell.
- parentTableCellOrNullObject — Gets the table cell that contains the field. If it isn't contained in a table cell, then this method will return an object with its isNullObject property set to true. For further information, see “OrNullObject methods and properties”: https://learn.microsoft.com/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties.
- parentTableOrNullObject — Gets the table that contains the field. If it isn't contained in a table, then this method will return an object with its isNullObject property set to true. For further information, see “OrNullObject methods and properties”: https://learn.microsoft.com/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties.
- result — Gets the field's result data.
- showCodes — Specifies whether the field codes are displayed for the specified field. true if the field codes are displayed, false otherwise.
- type — Gets the field's type.

## Property Details

### $all
Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).

```typescript
$all?: boolean;
```

Property value: boolean

---

### code
Specifies the field's code instruction.

```typescript
code?: boolean;
```

Property value: boolean

Remarks
- [API set: WordApi 1.4]
- Note: The ability to set the code was introduced in WordApi 1.5.

---

### data
Specifies data in an "Addin" field. If the field isn't an "Addin" field, it is null and it will throw a general exception when code attempts to set it.

```typescript
data?: boolean;
```

Property value: boolean

Remarks
- [API set: WordApi 1.5]

---

### kind
Gets the field's kind.

```typescript
kind?: boolean;
```

Property value: boolean

Remarks
- [API set: WordApi 1.5]

---

### linkFormat
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets a LinkFormat object that represents the link options of the field.

```typescript
linkFormat?: Word.Interfaces.LinkFormatLoadOptions;
```

Property value: Word.Interfaces.LinkFormatLoadOptions (https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.linkformatloadoptions)

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)]

---

### locked
Specifies whether the field is locked. true if the field is locked, false otherwise.

```typescript
locked?: boolean;
```

Property value: boolean

Remarks
- [API set: WordApi 1.5]

---

### oleFormat
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets an OleFormat object that represents the OLE characteristics (other than linking) for the field.

```typescript
oleFormat?: Word.Interfaces.OleFormatLoadOptions;
```

Property value: Word.Interfaces.OleFormatLoadOptions (https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.oleformatloadoptions)

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)]

---

### parentBody
Gets the parent body of the field.

```typescript
parentBody?: Word.Interfaces.BodyLoadOptions;
```

Property value: Word.Interfaces.BodyLoadOptions (https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.bodyloadoptions)

Remarks
- [API set: WordApi 1.4]

---

### parentContentControl
Gets the content control that contains the field. Throws an ItemNotFound error if there isn't a parent content control.

```typescript
parentContentControl?: Word.Interfaces.ContentControlLoadOptions;
```

Property value: Word.Interfaces.ContentControlLoadOptions (https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.contentcontrolloadoptions)

Remarks
- [API set: WordApi 1.4]

---

### parentContentControlOrNullObject
Gets the content control that contains the field. If there isn't a parent content control, then this method will return an object with its isNullObject property set to true. For further information, see “OrNullObject methods and properties”: https://learn.microsoft.com/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties.

```typescript
parentContentControlOrNullObject?: Word.Interfaces.ContentControlLoadOptions;
```

Property value: Word.Interfaces.ContentControlLoadOptions (https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.contentcontrolloadoptions)

Remarks
- [API set: WordApi 1.4]

---

### parentTable
Gets the table that contains the field. Throws an ItemNotFound error if it isn't contained in a table.

```typescript
parentTable?: Word.Interfaces.TableLoadOptions;
```

Property value: Word.Interfaces.TableLoadOptions (https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.tableloadoptions)

Remarks
- [API set: WordApi 1.4]

---

### parentTableCell
Gets the table cell that contains the field. Throws an ItemNotFound error if it isn't contained in a table cell.

```typescript
parentTableCell?: Word.Interfaces.TableCellLoadOptions;
```

Property value: Word.Interfaces.TableCellLoadOptions (https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.tablecellloadoptions)

Remarks
- [API set: WordApi 1.4]

---

### parentTableCellOrNullObject
Gets the table cell that contains the field. If it isn't contained in a table cell, then this method will return an object with its isNullObject property set to true. For further information, see “OrNullObject methods and properties”: https://learn.microsoft.com/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties.

```typescript
parentTableCellOrNullObject?: Word.Interfaces.TableCellLoadOptions;
```

Property value: Word.Interfaces.TableCellLoadOptions (https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.tablecellloadoptions)

Remarks
- [API set: WordApi 1.4]

---

### parentTableOrNullObject
Gets the table that contains the field. If it isn't contained in a table, then this method will return an object with its isNullObject property set to true. For further information, see “OrNullObject methods and properties”: https://learn.microsoft.com/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties.

```typescript
parentTableOrNullObject?: Word.Interfaces.TableLoadOptions;
```

Property value: Word.Interfaces.TableLoadOptions (https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.tableloadoptions)

Remarks
- [API set: WordApi 1.4]

---

### result
Gets the field's result data.

```typescript
result?: Word.Interfaces.RangeLoadOptions;
```

Property value: Word.Interfaces.RangeLoadOptions (https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.rangeloadoptions)

Remarks
- [API set: WordApi 1.4]

---

### showCodes
Specifies whether the field codes are displayed for the specified field. true if the field codes are displayed, false otherwise.

```typescript
showCodes?: boolean;
```

Property value: boolean

Remarks
- [API set: WordApiDesktop 1.1]

---

### type
Gets the field's type.

```typescript
type?: boolean;
```

Property value: boolean

Remarks
- [API set: WordApi 1.5]