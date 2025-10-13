# Word.Interfaces.FieldCollectionLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Contains a collection of [Word.Field](/en-us/javascript/api/word/word.field) objects.

## Remarks

[ API set: WordApi 1.4 ]

Important: To learn more about which fields can be inserted, see the Word.Range.insertField API introduced in requirement set 1.5. Support for managing fields is similar to what's available in the Word UI. However, the Word UI on the web primarily only supports fields as read-only (see [Field codes in Word for the web](https://support.microsoft.com/office/d8f46094-13c3-4966-98c3-259748f3caf1)). To learn more about Word UI clients that more fully support fields, see the product list at the beginning of [Insert, edit, and view fields in Word](https://support.microsoft.com/office/c429bbb0-8669-48a7-bd24-bab6ba6b06bb).

## Properties

- $all  
  Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).

- code  
  For EACH ITEM in the collection: Specifies the field's code instruction.

- data  
  For EACH ITEM in the collection: Specifies data in an "Addin" field. If the field isn't an "Addin" field, it is null and it will throw a general exception when code attempts to set it.

- kind  
  For EACH ITEM in the collection: Gets the field's kind.

- linkFormat  
  For EACH ITEM in the collection: Gets a LinkFormat object that represents the link options of the field.

- locked  
  For EACH ITEM in the collection: Specifies whether the field is locked. true if the field is locked, false otherwise.

- oleFormat  
  For EACH ITEM in the collection: Gets an OleFormat object that represents the OLE characteristics (other than linking) for the field.

- parentBody  
  For EACH ITEM in the collection: Gets the parent body of the field.

- parentContentControl  
  For EACH ITEM in the collection: Gets the content control that contains the field. Throws an ItemNotFound error if there isn't a parent content control.

- parentContentControlOrNullObject  
  For EACH ITEM in the collection: Gets the content control that contains the field. If there isn't a parent content control, then this method will return an object with its isNullObject property set to true. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

- parentTable  
  For EACH ITEM in the collection: Gets the table that contains the field. Throws an ItemNotFound error if it isn't contained in a table.

- parentTableCell  
  For EACH ITEM in the collection: Gets the table cell that contains the field. Throws an ItemNotFound error if it isn't contained in a table cell.

- parentTableCellOrNullObject  
  For EACH ITEM in the collection: Gets the table cell that contains the field. If it isn't contained in a table cell, then this method will return an object with its isNullObject property set to true. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

- parentTableOrNullObject  
  For EACH ITEM in the collection: Gets the table that contains the field. If it isn't contained in a table, then this method will return an object with its isNullObject property set to true. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

- result  
  For EACH ITEM in the collection: Gets the field's result data.

- showCodes  
  For EACH ITEM in the collection: Specifies whether the field codes are displayed for the specified field. true if the field codes are displayed, false otherwise.

- type  
  For EACH ITEM in the collection: Gets the field's type.

## Property Details

### $all

Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).

```typescript
$all?: boolean;
```

Property Value: boolean

---

### code

For EACH ITEM in the collection: Specifies the field's code instruction.

```typescript
code?: boolean;
```

Property Value: boolean

Remarks  
[ API set: WordApi 1.4 ]  
Note: The ability to set the code was introduced in WordApi 1.5.

---

### data

For EACH ITEM in the collection: Specifies data in an "Addin" field. If the field isn't an "Addin" field, it is null and it will throw a general exception when code attempts to set it.

```typescript
data?: boolean;
```

Property Value: boolean

Remarks  
[ API set: WordApi 1.5 ]

---

### kind

For EACH ITEM in the collection: Gets the field's kind.

```typescript
kind?: boolean;
```

Property Value: boolean

Remarks  
[ API set: WordApi 1.5 ]

---

### linkFormat

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Gets a LinkFormat object that represents the link options of the field.

```typescript
linkFormat?: Word.Interfaces.LinkFormatLoadOptions;
```

Property Value: [Word.Interfaces.LinkFormatLoadOptions](/en-us/javascript/api/word/word.interfaces.linkformatloadoptions)

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### locked

For EACH ITEM in the collection: Specifies whether the field is locked. true if the field is locked, false otherwise.

```typescript
locked?: boolean;
```

Property Value: boolean

Remarks  
[ API set: WordApi 1.5 ]

---

### oleFormat

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Gets an OleFormat object that represents the OLE characteristics (other than linking) for the field.

```typescript
oleFormat?: Word.Interfaces.OleFormatLoadOptions;
```

Property Value: [Word.Interfaces.OleFormatLoadOptions](/en-us/javascript/api/word/word.interfaces.oleformatloadoptions)

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### parentBody

For EACH ITEM in the collection: Gets the parent body of the field.

```typescript
parentBody?: Word.Interfaces.BodyLoadOptions;
```

Property Value: [Word.Interfaces.BodyLoadOptions](/en-us/javascript/api/word/word.interfaces.bodyloadoptions)

Remarks  
[ API set: WordApi 1.4 ]

---

### parentContentControl

For EACH ITEM in the collection: Gets the content control that contains the field. Throws an ItemNotFound error if there isn't a parent content control.

```typescript
parentContentControl?: Word.Interfaces.ContentControlLoadOptions;
```

Property Value: [Word.Interfaces.ContentControlLoadOptions](/en-us/javascript/api/word/word.interfaces.contentcontrolloadoptions)

Remarks  
[ API set: WordApi 1.4 ]

---

### parentContentControlOrNullObject

For EACH ITEM in the collection: Gets the content control that contains the field. If there isn't a parent content control, then this method will return an object with its isNullObject property set to true. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
parentContentControlOrNullObject?: Word.Interfaces.ContentControlLoadOptions;
```

Property Value: [Word.Interfaces.ContentControlLoadOptions](/en-us/javascript/api/word/word.interfaces.contentcontrolloadoptions)

Remarks  
[ API set: WordApi 1.4 ]

---

### parentTable

For EACH ITEM in the collection: Gets the table that contains the field. Throws an ItemNotFound error if it isn't contained in a table.

```typescript
parentTable?: Word.Interfaces.TableLoadOptions;
```

Property Value: [Word.Interfaces.TableLoadOptions](/en-us/javascript/api/word/word.interfaces.tableloadoptions)

Remarks  
[ API set: WordApi 1.4 ]

---

### parentTableCell

For EACH ITEM in the collection: Gets the table cell that contains the field. Throws an ItemNotFound error if it isn't contained in a table cell.

```typescript
parentTableCell?: Word.Interfaces.TableCellLoadOptions;
```

Property Value: [Word.Interfaces.TableCellLoadOptions](/en-us/javascript/api/word/word.interfaces.tablecellloadoptions)

Remarks  
[ API set: WordApi 1.4 ]

---

### parentTableCellOrNullObject

For EACH ITEM in the collection: Gets the table cell that contains the field. If it isn't contained in a table cell, then this method will return an object with its isNullObject property set to true. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
parentTableCellOrNullObject?: Word.Interfaces.TableCellLoadOptions;
```

Property Value: [Word.Interfaces.TableCellLoadOptions](/en-us/javascript/api/word/word.interfaces.tablecellloadoptions)

Remarks  
[ API set: WordApi 1.4 ]

---

### parentTableOrNullObject

For EACH ITEM in the collection: Gets the table that contains the field. If it isn't contained in a table, then this method will return an object with its isNullObject property set to true. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
parentTableOrNullObject?: Word.Interfaces.TableLoadOptions;
```

Property Value: [Word.Interfaces.TableLoadOptions](/en-us/javascript/api/word/word.interfaces.tableloadoptions)

Remarks  
[ API set: WordApi 1.4 ]

---

### result

For EACH ITEM in the collection: Gets the field's result data.

```typescript
result?: Word.Interfaces.RangeLoadOptions;
```

Property Value: [Word.Interfaces.RangeLoadOptions](/en-us/javascript/api/word/word.interfaces.rangeloadoptions)

Remarks  
[ API set: WordApi 1.4 ]

---

### showCodes

For EACH ITEM in the collection: Specifies whether the field codes are displayed for the specified field. true if the field codes are displayed, false otherwise.

```typescript
showCodes?: boolean;
```

Property Value: boolean

Remarks  
[ API set: WordApiDesktop 1.1 ]

---

### type

For EACH ITEM in the collection: Gets the field's type.

```typescript
type?: boolean;
```

Property Value: boolean

Remarks  
[ API set: WordApi 1.5 ]