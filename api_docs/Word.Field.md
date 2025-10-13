# Word.Field class

Package: [word](/en-us/javascript/api/word)

Represents a field.

Extends
- [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks
[ API set: WordApi 1.4 ]

Important: To learn more about which fields can be inserted, see the Word.Range.insertField API introduced in requirement set 1.5. Support for managing fields is similar to what's available in the Word UI. However, the Word UI on the web primarily only supports fields as read-only (see Field codes in Word for the web). To learn more about Word UI clients that more fully support fields, see the product list at the beginning of Insert, edit, and view fields in Word.

Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-fields.yaml

// Gets the first field in the document.
await Word.run(async (context) => {
  const field: Word.Field = context.document.body.fields.getFirstOrNullObject();
  field.load(["code", "result", "locked", "type", "data", "kind"]);

  await context.sync();

  if (field.isNullObject) {
    console.log("This document has no fields.");
  } else {
    console.log("Code of first field: " + field.code, "Result of first field: " + JSON.stringify(field.result), "Type of first field: " + field.type, "Is the first field locked? " + field.locked, "Kind of the first field: " + field.kind);
  }
});
```

## Properties
- code — Specifies the field's code instruction.
- context — The request context associated with the object. This connects the add-in's process to the Office host application's process.
- data — Specifies data in an "Addin" field. If the field isn't an "Addin" field, it is null and it will throw a general exception when code attempts to set it.
- kind — Gets the field's kind.
- linkFormat — Gets a LinkFormat object that represents the link options of the field.
- locked — Specifies whether the field is locked. true if the field is locked, false otherwise.
- oleFormat — Gets an OleFormat object that represents the OLE characteristics (other than linking) for the field.
- parentBody — Gets the parent body of the field.
- parentContentControl — Gets the content control that contains the field. Throws an ItemNotFound error if there isn't a parent content control.
- parentContentControlOrNullObject — Gets the content control that contains the field. If there isn't a parent content control, then this method will return an object with its isNullObject property set to true. For further information, see *OrNullObject methods and properties.
- parentTable — Gets the table that contains the field. Throws an ItemNotFound error if it isn't contained in a table.
- parentTableCell — Gets the table cell that contains the field. Throws an ItemNotFound error if it isn't contained in a table cell.
- parentTableCellOrNullObject — Gets the table cell that contains the field. If it isn't contained in a table cell, then this method will return an object with its isNullObject property set to true. For further information, see *OrNullObject methods and properties.
- parentTableOrNullObject — Gets the table that contains the field. If it isn't contained in a table, then this method will return an object with its isNullObject property set to true. For further information, see *OrNullObject methods and properties.
- result — Gets the field's result data.
- showCodes — Specifies whether the field codes are displayed for the specified field. true if the field codes are displayed, false otherwise.
- type — Gets the field's type.

## Methods
- copyToClipboard() — Copies the field to the Clipboard.
- cut() — Removes the field from the document and places it on the Clipboard.
- delete() — Deletes the field.
- doClick() — Clicks the field.
- getNext() — Gets the next field. Throws an ItemNotFound error if this field is the last one.
- getNextOrNullObject() — Gets the next field. If this field is the last one, then this method will return an object with its isNullObject property set to true. For further information, see *OrNullObject methods and properties.
- load(options) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNames) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNamesAndPaths) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- select(selectionMode) — Selects the field.
- select(selectionMode) — Selects the field.
- set(properties, options) — Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
- set(properties) — Sets multiple properties on the object at the same time, based on an existing loaded object.
- toJSON() — Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.Field object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.FieldData) that contains shallow copies of any loaded child properties from the original object.
- track() — Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
- unlink() — Replaces the field with its most recent result.
- untrack() — Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.
- updateResult() — Updates the field.
- updateSource() — Saves the changes made to the results of an INCLUDETEXT field back to the source document.

## Property Details

### code
Specifies the field's code instruction.

```typescript
code: string;
```

Property Value
- string

Remarks
[ API set: WordApi 1.4 ]

Note: The ability to set the code was introduced in WordApi 1.5.

Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-fields.yaml

// Gets the first field in the document.
await Word.run(async (context) => {
  const field: Word.Field = context.document.body.fields.getFirstOrNullObject();
  field.load(["code", "result", "locked", "type", "data", "kind"]);

  await context.sync();

  if (field.isNullObject) {
    console.log("This document has no fields.");
  } else {
    console.log("Code of first field: " + field.code, "Result of first field: " + JSON.stringify(field.result), "Type of first field: " + field.type, "Is the first field locked? " + field.locked, "Kind of the first field: " + field.kind);
  }
});
```

### context
The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value
- [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### data
Specifies data in an "Addin" field. If the field isn't an "Addin" field, it is null and it will throw a general exception when code attempts to set it.

```typescript
data: string;
```

Property Value
- string

Remarks
[ API set: WordApi 1.5 ]

### kind
Gets the field's kind.

```typescript
readonly kind: Word.FieldKind | "None" | "Hot" | "Warm" | "Cold";
```

Property Value
- [Word.FieldKind](/en-us/javascript/api/word/word.fieldkind) | "None" | "Hot" | "Warm" | "Cold"

Remarks
[ API set: WordApi 1.5 ]

Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-fields.yaml

// Gets the first field in the document.
await Word.run(async (context) => {
  const field: Word.Field = context.document.body.fields.getFirstOrNullObject();
  field.load(["code", "result", "locked", "type", "data", "kind"]);

  await context.sync();

  if (field.isNullObject) {
    console.log("This document has no fields.");
  } else {
    console.log("Code of first field: " + field.code, "Result of first field: " + JSON.stringify(field.result), "Type of first field: " + field.type, "Is the first field locked? " + field.locked, "Kind of the first field: " + field.kind);
  }
});
```

### linkFormat
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets a LinkFormat object that represents the link options of the field.

```typescript
readonly linkFormat: Word.LinkFormat;
```

Property Value
- [Word.LinkFormat](/en-us/javascript/api/word/word.linkformat)

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### locked
Specifies whether the field is locked. true if the field is locked, false otherwise.

```typescript
locked: boolean;
```

Property Value
- boolean

Remarks
[ API set: WordApi 1.5 ]

Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-fields.yaml

// Gets the first field in the selection and toggles between setting it to locked or unlocked.
await Word.run(async (context) => {
  let field = context.document.getSelection().fields.getFirstOrNullObject();
  field.load(["code", "result", "type", "locked"]);
  await context.sync();

  if (field.isNullObject) {
    console.log("The selection has no fields.");
  } else {
    console.log(`The first field in the selection is currently ${field.locked ? "locked" : "unlocked"}.`);
    field.locked = !field.locked;
    await context.sync();

    console.log(`The first field in the selection is now ${field.locked ? "locked" : "unlocked"}.`);
  }
});
```

### oleFormat
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets an OleFormat object that represents the OLE characteristics (other than linking) for the field.

```typescript
readonly oleFormat: Word.OleFormat;
```

Property Value
- [Word.OleFormat](/en-us/javascript/api/word/word.oleformat)

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### parentBody
Gets the parent body of the field.

```typescript
readonly parentBody: Word.Body;
```

Property Value
- [Word.Body](/en-us/javascript/api/word/word.body)

Remarks
[ API set: WordApi 1.4 ]

Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-fields.yaml

// Gets the parent body of the first field in the document.
await Word.run(async (context) => {
  const field: Word.Field = context.document.body.fields.getFirstOrNullObject();
  field.load("parentBody/text");

  await context.sync();

  if (field.isNullObject) {
    console.log("This document has no fields.");
  } else {
    const parentBody: Word.Body = field.parentBody;
    console.log("Text of first field's parent body: " + JSON.stringify(parentBody.text));
  }
});
```

### parentContentControl
Gets the content control that contains the field. Throws an ItemNotFound error if there isn't a parent content control.

```typescript
readonly parentContentControl: Word.ContentControl;
```

Property Value
- [Word.ContentControl](/en-us/javascript/api/word/word.contentcontrol)

Remarks
[ API set: WordApi 1.4 ]

### parentContentControlOrNullObject
Gets the content control that contains the field. If there isn't a parent content control, then this method will return an object with its isNullObject property set to true. For further information, see *OrNullObject methods and properties.

```typescript
readonly parentContentControlOrNullObject: Word.ContentControl;
```

Property Value
- [Word.ContentControl](/en-us/javascript/api/word/word.contentcontrol)

Remarks
[ API set: WordApi 1.4 ]

### parentTable
Gets the table that contains the field. Throws an ItemNotFound error if it isn't contained in a table.

```typescript
readonly parentTable: Word.Table;
```

Property Value
- [Word.Table](/en-us/javascript/api/word/word.table)

Remarks
[ API set: WordApi 1.4 ]

### parentTableCell
Gets the table cell that contains the field. Throws an ItemNotFound error if it isn't contained in a table cell.

```typescript
readonly parentTableCell: Word.TableCell;
```

Property Value
- [Word.TableCell](/en-us/javascript/api/word/word.tablecell)

Remarks
[ API set: WordApi 1.4 ]

### parentTableCellOrNullObject
Gets the table cell that contains the field. If it isn't contained in a table cell, then this method will return an object with its isNullObject property set to true. For further information, see *OrNullObject methods and properties.

```typescript
readonly parentTableCellOrNullObject: Word.TableCell;
```

Property Value
- [Word.TableCell](/en-us/javascript/api/word/word.tablecell)

Remarks
[ API set: WordApi 1.4 ]

### parentTableOrNullObject
Gets the table that contains the field. If it isn't contained in a table, then this method will return an object with its isNullObject property set to true. For further information, see *OrNullObject methods and properties.

```typescript
readonly parentTableOrNullObject: Word.Table;
```

Property Value
- [Word.Table](/en-us/javascript/api/word/word.table)

Remarks
[ API set: WordApi 1.4 ]

### result
Gets the field's result data.

```typescript
readonly result: Word.Range;
```

Property Value
- [Word.Range](/en-us/javascript/api/word/word.range)

Remarks
[ API set: WordApi 1.4 ]

Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-fields.yaml

// Gets the first field in the document.
await Word.run(async (context) => {
  const field: Word.Field = context.document.body.fields.getFirstOrNullObject();
  field.load(["code", "result", "locked", "type", "data", "kind"]);

  await context.sync();

  if (field.isNullObject) {
    console.log("This document has no fields.");
  } else {
    console.log("Code of first field: " + field.code, "Result of first field: " + JSON.stringify(field.result), "Type of first field: " + field.type, "Is the first field locked? " + field.locked, "Kind of the first field: " + field.kind);
  }
});
```

### showCodes
Specifies whether the field codes are displayed for the specified field. true if the field codes are displayed, false otherwise.

```typescript
showCodes: boolean;
```

Property Value
- boolean

Remarks
[ API set: WordApiDesktop 1.1 ]

### type
Gets the field's type.

```typescript
readonly type: Word.FieldType | "Addin" | "AddressBlock" | "Advance" | "Ask" | "Author" | "AutoText" | "AutoTextList" | "BarCode" | "Bibliography" | "BidiOutline" | "Citation" | "Comments" | "Compare" | "CreateDate" | "Data" | "Database" | "Date" | "DisplayBarcode" | "DocProperty" | "DocVariable" | "EditTime" | "Embedded" | "EQ" | "Expression" | "FileName" | "FileSize" | "FillIn" | "FormCheckbox" | "FormDropdown" | "FormText" | "GotoButton" | "GreetingLine" | "Hyperlink" | "If" | "Import" | "Include" | "IncludePicture" | "IncludeText" | "Index" | "Info" | "Keywords" | "LastSavedBy" | "Link" | "ListNum" | "MacroButton" | "MergeBarcode" | "MergeField" | "MergeRec" | "MergeSeq" | "Next" | "NextIf" | "NoteRef" | "NumChars" | "NumPages" | "NumWords" | "OCX" | "Page" | "PageRef" | "Print" | "PrintDate" | "Private" | "Quote" | "RD" | "Ref" | "RevNum" | "SaveDate" | "Section" | "SectionPages" | "Seq" | "Set" | "Shape" | "SkipIf" | "StyleRef" | "Subject" | "Subscriber" | "Symbol" | "TA" | "TC" | "Template" | "Time" | "Title" | "TOA" | "TOC" | "UserAddress" | "UserInitials" | "UserName" | "XE" | "Empty" | "Others" | "Undefined";
```

Property Value
- [Word.FieldType](/en-us/javascript/api/word/word.fieldtype) | "Addin" | "AddressBlock" | "Advance" | "Ask" | "Author" | "AutoText" | "AutoTextList" | "BarCode" | "Bibliography" | "BidiOutline" | "Citation" | "Comments" | "Compare" | "CreateDate" | "Data" | "Database" | "Date" | "DisplayBarcode" | "DocProperty" | "DocVariable" | "EditTime" | "Embedded" | "EQ" | "Expression" | "FileName" | "FileSize" | "FillIn" | "FormCheckbox" | "FormDropdown" | "FormText" | "GotoButton" | "GreetingLine" | "Hyperlink" | "If" | "Import" | "Include" | "IncludePicture" | "IncludeText" | "Index" | "Info" | "Keywords" | "LastSavedBy" | "Link" | "ListNum" | "MacroButton" | "MergeBarcode" | "MergeField" | "MergeRec" | "MergeSeq" | "Next" | "NextIf" | "NoteRef" | "NumChars" | "NumPages" | "NumWords" | "OCX" | "Page" | "PageRef" | "Print" | "PrintDate" | "Private" | "Quote" | "RD" | "Ref" | "RevNum" | "SaveDate" | "Section" | "SectionPages" | "Seq" | "Set" | "Shape" | "SkipIf" | "StyleRef" | "Subject" | "Subscriber" | "Symbol" | "TA" | "TC" | "Template" | "Time" | "Title" | "TOA" | "TOC" | "UserAddress" | "UserInitials" | "UserName" | "XE" | "Empty" | "Others" | "Undefined"

Remarks
[ API set: WordApi 1.5 ]

Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-fields.yaml

// Gets the first field in the document.
await Word.run(async (context) => {
  const field: Word.Field = context.document.body.fields.getFirstOrNullObject();
  field.load(["code", "result", "locked", "type", "data", "kind"]);

  await context.sync();

  if (field.isNullObject) {
    console.log("This document has no fields.");
  } else {
    console.log("Code of first field: " + field.code, "Result of first field: " + JSON.stringify(field.result), "Type of first field: " + field.type, "Is the first field locked? " + field.locked, "Kind of the first field: " + field.kind);
  }
});
```

## Method Details

### copyToClipboard()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Copies the field to the Clipboard.

```typescript
copyToClipboard(): void;
```

Returns
- void

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### cut()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Removes the field from the document and places it on the Clipboard.

```typescript
cut(): void;
```

Returns
- void

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### delete()
Deletes the field.

```typescript
delete(): void;
```

Returns
- void

Remarks
[ API set: WordApi 1.5 ]

Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-fields.yaml

// Deletes the first field in the document.
await Word.run(async (context) => {
  const field: Word.Field = context.document.body.fields.getFirstOrNullObject();
  field.load();

  await context.sync();

  if (field.isNullObject) {
    console.log("This document has no fields.");
  } else {
    field.delete();
    await context.sync();

    console.log("The first field in the document was deleted.");
  }
});
```

### doClick()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Clicks the field.

```typescript
doClick(): void;
```

Returns
- void

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### getNext()
Gets the next field. Throws an ItemNotFound error if this field is the last one.

```typescript
getNext(): Word.Field;
```

Returns
- [Word.Field](/en-us/javascript/api/word/word.field)

Remarks
[ API set: WordApi 1.4 ]

### getNextOrNullObject()
Gets the next field. If this field is the last one, then this method will return an object with its isNullObject property set to true. For further information, see *OrNullObject methods and properties.

```typescript
getNextOrNullObject(): Word.Field;
```

Returns
- [Word.Field](/en-us/javascript/api/word/word.field)

Remarks
[ API set: WordApi 1.4 ]

### load(options)
Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(options?: Word.Interfaces.FieldLoadOptions): Word.Field;
```

Parameters
- options: [Word.Interfaces.FieldLoadOptions](/en-us/javascript/api/word/word.interfaces.fieldloadoptions)  
  Provides options for which properties of the object to load.

Returns
- [Word.Field](/en-us/javascript/api/word/word.field)

### load(propertyNames)
Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.Field;
```

Parameters
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns
- [Word.Field](/en-us/javascript/api/word/word.field)

### load(propertyNamesAndPaths)
Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.Field;
```

Parameters
- propertyNamesAndPaths:  
  {
  select?: string;
  expand?: string;
  }  
  propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

Returns
- [Word.Field](/en-us/javascript/api/word/word.field)

### select(selectionMode)
Selects the field.

```typescript
select(selectionMode?: Word.SelectionMode): void;
```

Parameters
- selectionMode: [Word.SelectionMode](/en-us/javascript/api/word/word.selectionmode)  
  Optional. The selection mode must be 'Select', 'Start', or 'End'. 'Select' is the default.

Returns
- void

Remarks
[ API set: WordApi 1.5 ]

Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-fields.yaml

// Gets and updates the first field in the selection.
await Word.run(async (context) => {
  let field = context.document.getSelection().fields.getFirstOrNullObject();
  field.load(["code", "result", "type", "locked"]);

  await context.sync();

  if (field.isNullObject) {
    console.log("No field in selection.");
  } else {
    console.log("Before updating:", "Code of first field: " + field.code, "Result of first field: " + JSON.stringify(field.result));

    field.updateResult();
    field.select();
    await context.sync();

    field.load(["code", "result"]);
    await context.sync();

    console.log("After updating:", "Code of first field: " + field.code, "Result of first field: " + JSON.stringify(field.result));
  }
});
```

### select(selectionMode)
Selects the field.

```typescript
select(selectionMode?: "Select" | "Start" | "End"): void;
```

Parameters
- selectionMode: "Select" | "Start" | "End"  
  Optional. The selection mode must be 'Select', 'Start', or 'End'. 'Select' is the default.

Returns
- void

Remarks
[ API set: WordApi 1.5 ]

### set(properties, options)
Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.FieldUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

Parameters
- properties: [Word.Interfaces.FieldUpdateData](/en-us/javascript/api/word/word.interfaces.fieldupdatedata)  
  A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
- options: [OfficeExtension.UpdateOptions](/en-us/javascript/api/office/officeextension.updateoptions)  
  Provides an option to suppress errors if the properties object tries to set any read-only properties.

Returns
- void

### set(properties)
Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.Field): void;
```

Parameters
- properties: [Word.Field](/en-us/javascript/api/word/word.field)

Returns
- void

### toJSON()
Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.Field object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.FieldData) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.FieldData;
```

Returns
- [Word.Interfaces.FieldData](/en-us/javascript/api/word/word.interfaces.fielddata)

### track()
Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.Field;
```

Returns
- [Word.Field](/en-us/javascript/api/word/word.field)

### unlink()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Replaces the field with its most recent result.

```typescript
unlink(): void;
```

Returns
- void

Remarks
[ API set: WordApi BETA (P