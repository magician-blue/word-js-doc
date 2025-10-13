# Word.Interfaces.FieldData interface

Package: [word](/en-us/javascript/api/word)

An interface describing the data returned by calling `field.toJSON()`.

## Properties

- [code](#code): Specifies the field's code instruction.
- [data](#data): Specifies data in an "Addin" field. If the field isn't an "Addin" field, it is `null` and it will throw a general exception when code attempts to set it.
- [kind](#kind): Gets the field's kind.
- [locked](#locked): Specifies whether the field is locked. `true` if the field is locked, `false` otherwise.
- [result](#result): Gets the field's result data.
- [showCodes](#showcodes): Specifies whether the field codes are displayed for the specified field. `true` if the field codes are displayed, `false` otherwise.
- [type](#type): Gets the field's type.

## Property Details

### code

Specifies the field's code instruction.

```typescript
code?: string;
```

Property Value
- string

Remarks  
[API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)  
Note: The ability to set the code was introduced in WordApi 1.5.

### data

Specifies data in an "Addin" field. If the field isn't an "Addin" field, it is `null` and it will throw a general exception when code attempts to set it.

```typescript
data?: string;
```

Property Value
- string

Remarks  
[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### kind

Gets the field's kind.

```typescript
kind?: Word.FieldKind | "None" | "Hot" | "Warm" | "Cold";
```

Property Value
- [Word.FieldKind](/en-us/javascript/api/word/word.fieldkind) | "None" | "Hot" | "Warm" | "Cold"

Remarks  
[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### locked

Specifies whether the field is locked. `true` if the field is locked, `false` otherwise.

```typescript
locked?: boolean;
```

Property Value
- boolean

Remarks  
[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### result

Gets the field's result data.

```typescript
result?: Word.Interfaces.RangeData;
```

Property Value
- [Word.Interfaces.RangeData](/en-us/javascript/api/word/word.interfaces.rangedata)

Remarks  
[API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### showCodes

Specifies whether the field codes are displayed for the specified field. `true` if the field codes are displayed, `false` otherwise.

```typescript
showCodes?: boolean;
```

Property Value
- boolean

Remarks  
[API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### type

Gets the field's type.

```typescript
type?: Word.FieldType | "Addin" | "AddressBlock" | "Advance" | "Ask" | "Author" | "AutoText" | "AutoTextList" | "BarCode" | "Bibliography" | "BidiOutline" | "Citation" | "Comments" | "Compare" | "CreateDate" | "Data" | "Database" | "Date" | "DisplayBarcode" | "DocProperty" | "DocVariable" | "EditTime" | "Embedded" | "EQ" | "Expression" | "FileName" | "FileSize" | "FillIn" | "FormCheckbox" | "FormDropdown" | "FormText" | "GotoButton" | "GreetingLine" | "Hyperlink" | "If" | "Import" | "Include" | "IncludePicture" | "IncludeText" | "Index" | "Info" | "Keywords" | "LastSavedBy" | "Link" | "ListNum" | "MacroButton" | "MergeBarcode" | "MergeField" | "MergeRec" | "MergeSeq" | "Next" | "NextIf" | "NoteRef" | "NumChars" | "NumPages" | "NumWords" | "OCX" | "Page" | "PageRef" | "Print" | "PrintDate" | "Private" | "Quote" | "RD" | "Ref" | "RevNum" | "SaveDate" | "Section" | "SectionPages" | "Seq" | "Set" | "Shape" | "SkipIf" | "StyleRef" | "Subject" | "Subscriber" | "Symbol" | "TA" | "TC" | "Template" | "Time" | "Title" | "TOA" | "TOC" | "UserAddress" | "UserInitials" | "UserName" | "XE" | "Empty" | "Others" | "Undefined";
```

Property Value
- [Word.FieldType](/en-us/javascript/api/word/word.fieldtype) | "Addin" | "AddressBlock" | "Advance" | "Ask" | "Author" | "AutoText" | "AutoTextList" | "BarCode" | "Bibliography" | "BidiOutline" | "Citation" | "Comments" | "Compare" | "CreateDate" | "Data" | "Database" | "Date" | "DisplayBarcode" | "DocProperty" | "DocVariable" | "EditTime" | "Embedded" | "EQ" | "Expression" | "FileName" | "FileSize" | "FillIn" | "FormCheckbox" | "FormDropdown" | "FormText" | "GotoButton" | "GreetingLine" | "Hyperlink" | "If" | "Import" | "Include" | "IncludePicture" | "IncludeText" | "Index" | "Info" | "Keywords" | "LastSavedBy" | "Link" | "ListNum" | "MacroButton" | "MergeBarcode" | "MergeField" | "MergeRec" | "MergeSeq" | "Next" | "NextIf" | "NoteRef" | "NumChars" | "NumPages" | "NumWords" | "OCX" | "Page" | "PageRef" | "Print" | "PrintDate" | "Private" | "Quote" | "RD" | "Ref" | "RevNum" | "SaveDate" | "Section" | "SectionPages" | "Seq" | "Set" | "Shape" | "SkipIf" | "StyleRef" | "Subject" | "Subscriber" | "Symbol" | "TA" | "TC" | "Template" | "Time" | "Title" | "TOA" | "TOC" | "UserAddress" | "UserInitials" | "UserName" | "XE" | "Empty" | "Others" | "Undefined"

Remarks  
[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)