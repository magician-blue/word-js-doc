# Word.Interfaces.ContentControlData interface

Package: [word](/en-us/javascript/api/word)

An interface describing the data returned by calling `contentControl.toJSON()`.

## Properties

- appearance  
  Specifies the appearance of the content control. The value can be 'BoundingBox', 'Tags', or 'Hidden'.

- buildingBlockGalleryContentControl  
  Gets the building block gallery-related data if the content control's [Word.ContentControlType](/en-us/javascript/api/word/word.contentcontroltype) is `BuildingBlockGallery`. It's `null` otherwise.

- cannotDelete  
  Specifies a value that indicates whether the user can delete the content control. Mutually exclusive with removeWhenEdited.

- cannotEdit  
  Specifies a value that indicates whether the user can edit the contents of the content control.

- checkboxContentControl  
  Gets the data of the content control when its type is `CheckBox`. It's `null` otherwise.

- color  
  Specifies the color of the content control. Color is specified in '#RRGGBB' format or by using the color name.

- comboBoxContentControl  
  Gets the data of the content control when its type is `ComboBox`. It's `null` otherwise.

- contentControls  
  Gets the collection of content control objects in the content control.

- datePickerContentControl  
  Gets the date picker-related data if the content control's [Word.ContentControlType](/en-us/javascript/api/word/word.contentcontroltype) is `DatePicker`. It's `null` otherwise.

- dropDownListContentControl  
  Gets the data of the content control when its type is `DropDownList`. It's `null` otherwise.

- fields  
  Gets the collection of field objects in the content control.

- font  
  Gets the text format of the content control. Use this to get and set font name, size, color, and other properties.

- groupContentControl  
  Gets the group-related data if the content control's [Word.ContentControlType](/en-us/javascript/api/word/word.contentcontroltype) is `Group`. It's `null` otherwise.

- id  
  Gets an integer that represents the content control identifier.

- inlinePictures  
  Gets the collection of InlinePicture objects in the content control. The collection doesn't include floating images.

- lists  
  Gets the collection of list objects in the content control.

- paragraphs  
  Gets the collection of paragraph objects in the content control.

- pictureContentControl  
  Gets the picture-related data if the content control's [Word.ContentControlType](/en-us/javascript/api/word/word.contentcontroltype) is `Picture`. It's `null` otherwise.

- placeholderText  
  Specifies the placeholder text of the content control. Dimmed text will be displayed when the content control is empty.

- removeWhenEdited  
  Specifies a value that indicates whether the content control is removed after it is edited. Mutually exclusive with cannotDelete.

- repeatingSectionContentControl  
  Gets the repeating section-related data if the content control's [Word.ContentControlType](/en-us/javascript/api/word/word.contentcontroltype) is `RepeatingSection`. It's `null` otherwise.

- style  
  Specifies the style name for the content control. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.

- styleBuiltIn  
  Specifies the built-in style name for the content control. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.

- subtype  
  Gets the content control subtype. The subtype can be 'RichTextInline', 'RichTextParagraphs', 'RichTextTableCell', 'RichTextTableRow' and 'RichTextTable' for rich text content controls, or 'PlainTextInline' and 'PlainTextParagraph' for plain text content controls, or 'CheckBox' for checkbox content controls.

- tables  
  Gets the collection of table objects in the content control.

- tag  
  Specifies a tag to identify a content control.

- text  
  Gets the text of the content control.

- title  
  Specifies the title for a content control.

- type  
  Gets the content control type. Only rich text, plain text, and checkbox content controls are supported currently.

- xmlMapping  
  Returns an `XmlMapping` object that represents the mapping of the content control to XML data in the data store of the document.

## Property Details

### appearance

Specifies the appearance of the content control. The value can be 'BoundingBox', 'Tags', or 'Hidden'.

```typescript
appearance?: Word.ContentControlAppearance | "BoundingBox" | "Tags" | "Hidden";
```

Property Value  
[Word.ContentControlAppearance](/en-us/javascript/api/word/word.contentcontrolappearance) | "BoundingBox" | "Tags" | "Hidden"

Remarks  
[API set: WordApi 1.1]

---

### buildingBlockGalleryContentControl

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the building block gallery-related data if the content control's [Word.ContentControlType](/en-us/javascript/api/word/word.contentcontroltype) is `BuildingBlockGallery`. It's `null` otherwise.

```typescript
buildingBlockGalleryContentControl?: Word.Interfaces.BuildingBlockGalleryContentControlData;
```

Property Value  
[Word.Interfaces.BuildingBlockGalleryContentControlData](/en-us/javascript/api/word/word.interfaces.buildingblockgallerycontentcontroldata)

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)]

---

### cannotDelete

Specifies a value that indicates whether the user can delete the content control. Mutually exclusive with removeWhenEdited.

```typescript
cannotDelete?: boolean;
```

Property Value  
boolean

Remarks  
[API set: WordApi 1.1]

---

### cannotEdit

Specifies a value that indicates whether the user can edit the contents of the content control.

```typescript
cannotEdit?: boolean;
```

Property Value  
boolean

Remarks  
[API set: WordApi 1.1]

---

### checkboxContentControl

Gets the data of the content control when its type is `CheckBox`. It's `null` otherwise.

```typescript
checkboxContentControl?: Word.Interfaces.CheckboxContentControlData;
```

Property Value  
[Word.Interfaces.CheckboxContentControlData](/en-us/javascript/api/word/word.interfaces.checkboxcontentcontroldata)

Remarks  
[API set: WordApi 1.7]

---

### color

Specifies the color of the content control. Color is specified in '#RRGGBB' format or by using the color name.

```typescript
color?: string;
```

Property Value  
string

Remarks  
[API set: WordApi 1.1]

---

### comboBoxContentControl

Gets the data of the content control when its type is `ComboBox`. It's `null` otherwise.

```typescript
comboBoxContentControl?: Word.Interfaces.ComboBoxContentControlData;
```

Property Value  
[Word.Interfaces.ComboBoxContentControlData](/en-us/javascript/api/word/word.interfaces.comboboxcontentcontroldata)

Remarks  
[API set: WordApi 1.9]

---

### contentControls

Gets the collection of content control objects in the content control.

```typescript
contentControls?: Word.Interfaces.ContentControlData[];
```

Property Value  
[Word.Interfaces.ContentControlData](/en-us/javascript/api/word/word.interfaces.contentcontroldata)[]

Remarks  
[API set: WordApi 1.1]

---

### datePickerContentControl

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the date picker-related data if the content control's [Word.ContentControlType](/en-us/javascript/api/word/word.contentcontroltype) is `DatePicker`. It's `null` otherwise.

```typescript
datePickerContentControl?: Word.Interfaces.DatePickerContentControlData;
```

Property Value  
[Word.Interfaces.DatePickerContentControlData](/en-us/javascript/api/word/word.interfaces.datepickercontentcontroldata)

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)]

---

### dropDownListContentControl

Gets the data of the content control when its type is `DropDownList`. It's `null` otherwise.

```typescript
dropDownListContentControl?: Word.Interfaces.DropDownListContentControlData;
```

Property Value  
[Word.Interfaces.DropDownListContentControlData](/en-us/javascript/api/word/word.interfaces.dropdownlistcontentcontroldata)

Remarks  
[API set: WordApi 1.9]

---

### fields

Gets the collection of field objects in the content control.

```typescript
fields?: Word.Interfaces.FieldData[];
```

Property Value  
[Word.Interfaces.FieldData](/en-us/javascript/api/word/word.interfaces.fielddata)[]

Remarks  
[API set: WordApi 1.4]

---

### font

Gets the text format of the content control. Use this to get and set font name, size, color, and other properties.

```typescript
font?: Word.Interfaces.FontData;
```

Property Value  
[Word.Interfaces.FontData](/en-us/javascript/api/word/word.interfaces.fontdata)

Remarks  
[API set: WordApi 1.1]

---

### groupContentControl

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the group-related data if the content control's [Word.ContentControlType](/en-us/javascript/api/word/word.contentcontroltype) is `Group`. It's `null` otherwise.

```typescript
groupContentControl?: Word.Interfaces.GroupContentControlData;
```

Property Value  
[Word.Interfaces.GroupContentControlData](/en-us/javascript/api/word/word.interfaces.groupcontentcontroldata)

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)]

---

### id

Gets an integer that represents the content control identifier.

```typescript
id?: number;
```

Property Value  
number

Remarks  
[API set: WordApi 1.1]

---

### inlinePictures

Gets the collection of InlinePicture objects in the content control. The collection doesn't include floating images.

```typescript
inlinePictures?: Word.Interfaces.InlinePictureData[];
```

Property Value  
[Word.Interfaces.InlinePictureData](/en-us/javascript/api/word/word.interfaces.inlinepicturedata)[]

Remarks  
[API set: WordApi 1.1]

---

### lists

Gets the collection of list objects in the content control.

```typescript
lists?: Word.Interfaces.ListData[];
```

Property Value  
[Word.Interfaces.ListData](/en-us/javascript/api/word/word.interfaces.listdata)[]

Remarks  
[API set: WordApi 1.3]

---

### paragraphs

Gets the collection of paragraph objects in the content control.

```typescript
paragraphs?: Word.Interfaces.ParagraphData[];
```

Property Value  
[Word.Interfaces.ParagraphData](/en-us/javascript/api/word/word.interfaces.paragraphdata)[]

Remarks  
[API set: WordApi 1.1]  
Important: For requirement sets 1.1 and 1.2, paragraphs in tables wholly contained within this content control aren't returned. From requirement set 1.3, paragraphs in such tables are also returned.

---

### pictureContentControl

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the picture-related data if the content control's [Word.ContentControlType](/en-us/javascript/api/word/word.contentcontroltype) is `Picture`. It's `null` otherwise.

```typescript
pictureContentControl?: Word.Interfaces.PictureContentControlData;
```

Property Value  
[Word.Interfaces.PictureContentControlData](/en-us/javascript/api/word/word.interfaces.picturecontentcontroldata)

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)]

---

### placeholderText

Specifies the placeholder text of the content control. Dimmed text will be displayed when the content control is empty.

```typescript
placeholderText?: string;
```

Property Value  
string

Remarks  
[API set: WordApi 1.1]

---

### removeWhenEdited

Specifies a value that indicates whether the content control is removed after it is edited. Mutually exclusive with cannotDelete.

```typescript
removeWhenEdited?: boolean;
```

Property Value  
boolean

Remarks  
[API set: WordApi 1.1]

---

### repeatingSectionContentControl

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the repeating section-related data if the content control's [Word.ContentControlType](/en-us/javascript/api/word/word.contentcontroltype) is `RepeatingSection`. It's `null` otherwise.

```typescript
repeatingSectionContentControl?: Word.Interfaces.RepeatingSectionContentControlData;
```

Property Value  
[Word.Interfaces.RepeatingSectionContentControlData](/en-us/javascript/api/word/word.interfaces.repeatingsectioncontentcontroldata)

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)]

---

### style

Specifies the style name for the content control. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.

```typescript
style?: string;
```

Property Value  
string

Remarks  
[API set: WordApi 1.1]

---

### styleBuiltIn

Specifies the built-in style name for the content control. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.

```typescript
styleBuiltIn?: Word.BuiltInStyleName | "Other" | "Normal" | "Heading1" | "Heading2" | "Heading3" | "Heading4" | "Heading5" | "Heading6" | "Heading7" | "Heading8" | "Heading9" | "Toc1" | "Toc2" | "Toc3" | "Toc4" | "Toc5" | "Toc6" | "Toc7" | "Toc8" | "Toc9" | "FootnoteText" | "Header" | "Footer" | "Caption" | "FootnoteReference" | "EndnoteReference" | "EndnoteText" | "Title" | "Subtitle" | "Hyperlink" | "Strong" | "Emphasis" | "NoSpacing" | "ListParagraph" | "Quote" | "IntenseQuote" | "SubtleEmphasis" | "IntenseEmphasis" | "SubtleReference" | "IntenseReference" | "BookTitle" | "Bibliography" | "TocHeading" | "TableGrid" | "PlainTable1" | "PlainTable2" | "PlainTable3" | "PlainTable4" | "PlainTable5" | "TableGridLight" | "GridTable1Light" | "GridTable1Light_Accent1" | "GridTable1Light_Accent2" | "GridTable1Light_Accent3" | "GridTable1Light_Accent4" | "GridTable1Light_Accent5" | "GridTable1Light_Accent6" | "GridTable2" | "GridTable2_Accent1" | "GridTable2_Accent2" | "GridTable2_Accent3" | "GridTable2_Accent4" | "GridTable2_Accent5" | "GridTable2_Accent6" | "GridTable3" | "GridTable3_Accent1" | "GridTable3_Accent2" | "GridTable3_Accent3" | "GridTable3_Accent4" | "GridTable3_Accent5" | "GridTable3_Accent6" | "GridTable4" | "GridTable4_Accent1" | "GridTable4_Accent2" | "GridTable4_Accent3" | "GridTable4_Accent4" | "GridTable4_Accent5" | "GridTable4_Accent6" | "GridTable5Dark" | "GridTable5Dark_Accent1" | "GridTable5Dark_Accent2" | "GridTable5Dark_Accent3" | "GridTable5Dark_Accent4" | "GridTable5Dark_Accent5" | "GridTable5Dark_Accent6" | "GridTable6Colorful" | "GridTable6Colorful_Accent1" | "GridTable6Colorful_Accent2" | "GridTable6Colorful_Accent3" | "GridTable6Colorful_Accent4" | "GridTable6Colorful_Accent5" | "GridTable6Colorful_Accent6" | "GridTable7Colorful" | "GridTable7Colorful_Accent1" | "GridTable7Colorful_Accent2" | "GridTable7Colorful_Accent3" | "GridTable7Colorful_Accent4" | "GridTable7Colorful_Accent5" | "GridTable7Colorful_Accent6" | "ListTable1Light" | "ListTable1Light_Accent1" | "ListTable1Light_Accent2" | "ListTable1Light_Accent3" | "ListTable1Light_Accent4" | "ListTable1Light_Accent5" | "ListTable1Light_Accent6" | "ListTable2" | "ListTable2_Accent1" | "ListTable2_Accent2" | "ListTable2_Accent3" | "ListTable2_Accent4" | "ListTable2_Accent5" | "ListTable2_Accent6" | "ListTable3" | "ListTable3_Accent1" | "ListTable3_Accent2" | "ListTable3_Accent3" | "ListTable3_Accent4" | "ListTable3_Accent5" | "ListTable3_Accent6" | "ListTable4" | "ListTable4_Accent1" | "ListTable4_Accent2" | "ListTable4_Accent3" | "ListTable4_Accent4" | "ListTable4_Accent5" | "ListTable4_Accent6" | "ListTable5Dark" | "ListTable5Dark_Accent1" | "ListTable5Dark_Accent2" | "ListTable5Dark_Accent3" | "ListTable5Dark_Accent4" | "ListTable5Dark_Accent5" | "ListTable5Dark_Accent6" | "ListTable6Colorful" | "ListTable6Colorful_Accent1" | "ListTable6Colorful_Accent2" | "ListTable6Colorful_Accent3" | "ListTable6Colorful_Accent4" | "ListTable6Colorful_Accent5" | "ListTable6Colorful_Accent6" | "ListTable7Colorful" | "ListTable7Colorful_Accent1" | "ListTable7Colorful_Accent2" | "ListTable7Colorful_Accent3" | "ListTable7Colorful_Accent4" | "ListTable7Colorful_Accent5" | "ListTable7Colorful_Accent6";
```

Property Value  
[Word.BuiltInStyleName](/en-us/javascript/api/word/word.builtinstylename) | "Other" | "Normal" | "Heading1" | "Heading2" | "Heading3" | "Heading4" | "Heading5" | "Heading6" | "Heading7" | "Heading8" | "Heading9" | "Toc1" | "Toc2" | "Toc3" | "Toc4" | "Toc5" | "Toc6" | "Toc7" | "Toc8" | "Toc9" | "FootnoteText" | "Header" | "Footer" | "Caption" | "FootnoteReference" | "EndnoteReference" | "EndnoteText" | "Title" | "Subtitle" | "Hyperlink" | "Strong" | "Emphasis" | "NoSpacing" | "ListParagraph" | "Quote" | "IntenseQuote" | "SubtleEmphasis" | "IntenseEmphasis" | "SubtleReference" | "IntenseReference" | "BookTitle" | "Bibliography" | "TocHeading" | "TableGrid" | "PlainTable1" | "PlainTable2" | "PlainTable3" | "PlainTable4" | "PlainTable5" | "TableGridLight" | "GridTable1Light" | "GridTable1Light_Accent1" | "GridTable1Light_Accent2" | "GridTable1Light_Accent3" | "GridTable1Light_Accent4" | "GridTable1Light_Accent5" | "GridTable1Light_Accent6" | "GridTable2" | "GridTable2_Accent1" | "GridTable2_Accent2" | "GridTable2_Accent3" | "GridTable2_Accent4" | "GridTable2_Accent5" | "GridTable2_Accent6" | "GridTable3" | "GridTable3_Accent1" | "GridTable3_Accent2" | "GridTable3_Accent3" | "GridTable3_Accent4" | "GridTable3_Accent5" | "GridTable3_Accent6" | "GridTable4" | "GridTable4_Accent1" | "GridTable4_Accent2" | "GridTable4_Accent3" | "GridTable4_Accent4" | "GridTable4_Accent5" | "GridTable4_Accent6" | "GridTable5Dark" | "GridTable5Dark_Accent1" | "GridTable5Dark_Accent2" | "GridTable5Dark_Accent3" | "GridTable5Dark_Accent4" | "GridTable5Dark_Accent5" | "GridTable5Dark_Accent6" | "GridTable6Colorful" | "GridTable6Colorful_Accent1" | "GridTable6Colorful_Accent2" | "GridTable6Colorful_Accent3" | "GridTable6Colorful_Accent4" | "GridTable6Colorful_Accent5" | "GridTable6Colorful_Accent6" | "GridTable7Colorful" | "GridTable7Colorful_Accent1" | "GridTable7Colorful_Accent2" | "GridTable7Colorful_Accent3" | "GridTable7Colorful_Accent4" | "GridTable7Colorful_Accent5" | "GridTable7Colorful_Accent6" | "ListTable1Light" | "ListTable1Light_Accent1" | "ListTable1Light_Accent2" | "ListTable1Light_Accent3" | "ListTable1Light_Accent4" | "ListTable1Light_Accent5" | "ListTable1Light_Accent6" | "ListTable2" | "ListTable2_Accent1" | "ListTable2_Accent2" | "ListTable2_Accent3" | "ListTable2_Accent4" | "ListTable2_Accent5" | "ListTable2_Accent6" | "ListTable3" | "ListTable3_Accent1" | "ListTable3_Accent2" | "ListTable3_Accent3" | "ListTable3_Accent4" | "ListTable3_Accent5" | "ListTable3_Accent6" | "ListTable4" | "ListTable4_Accent1" | "ListTable4_Accent2" | "ListTable4_Accent3" | "ListTable4_Accent4" | "ListTable4_Accent5" | "ListTable4_Accent6" | "ListTable5Dark" | "ListTable5Dark_Accent1" | "ListTable5Dark_Accent2" | "ListTable5Dark_Accent3" | "ListTable5Dark_Accent4" | "ListTable5Dark_Accent5" | "ListTable5Dark_Accent6" | "ListTable6Colorful" | "ListTable6Colorful_Accent1" | "ListTable6Colorful_Accent2" | "ListTable6Colorful_Accent3" | "ListTable6Colorful_Accent4" | "ListTable6Colorful_Accent5" | "ListTable6Colorful_Accent6" | "ListTable7Colorful" | "ListTable7Colorful_Accent1" | "ListTable7Colorful_Accent2" | "ListTable7Colorful_Accent3" | "ListTable7Colorful_Accent4" | "ListTable7Colorful_Accent5" | "ListTable7Colorful_Accent6"

Remarks  
[API set: WordApi 1.3]

---

### subtype

Gets the content control subtype. The subtype can be 'RichTextInline', 'RichTextParagraphs', 'RichTextTableCell', 'RichTextTableRow' and 'RichTextTable' for rich text content controls, or 'PlainTextInline' and 'PlainTextParagraph' for plain text content controls, or 'CheckBox' for checkbox content controls.

```typescript
subtype?: Word.ContentControlType | "Unknown" | "RichTextInline" | "RichTextParagraphs" | "RichTextTableCell" | "RichTextTableRow" | "RichTextTable" | "PlainTextInline" | "PlainTextParagraph" | "Picture" | "BuildingBlockGallery" | "CheckBox" | "ComboBox" | "DropDownList" | "DatePicker" | "RepeatingSection" | "RichText" | "PlainText" | "Group";
```

Property Value  
[Word.ContentControlType](/en-us/javascript/api/word/word.contentcontroltype) | "Unknown" | "RichTextInline" | "RichTextParagraphs" | "RichTextTableCell" | "RichTextTableRow" | "RichTextTable" | "PlainTextInline" | "PlainTextParagraph" | "Picture" | "BuildingBlockGallery" | "CheckBox" | "ComboBox" | "DropDownList" | "DatePicker" | "RepeatingSection" | "RichText" | "PlainText" | "Group"

Remarks  
[API set: WordApi 1.3]

---

### tables

Gets the collection of table objects in the content control.

```typescript
tables?: Word.Interfaces.TableData[];
```

Property Value  
[Word.Interfaces.TableData](/en-us/javascript/api/word/word.interfaces.tabledata)[]

Remarks  
[API set: WordApi 1.3]

---

### tag

Specifies a tag to identify a content control.

```typescript
tag?: string;
```

Property Value  
string

Remarks  
[API set: WordApi 1.1]

---

### text

Gets the text of the content control.

```typescript
text?: string;
```

Property Value  
string

Remarks  
[API set: WordApi 1.1]

---

### title

Specifies the title for a content control.

```typescript
title?: string;
```

Property Value  
string

Remarks  
[API set: WordApi 1.1]

---

### type

Gets the content control type. Only rich text, plain text, and checkbox content controls are supported currently.

```typescript
type?: Word.ContentControlType | "Unknown" | "RichTextInline" | "RichTextParagraphs" | "RichTextTableCell" | "RichTextTableRow" | "RichTextTable" | "PlainTextInline" | "PlainTextParagraph" | "Picture" | "BuildingBlockGallery" | "CheckBox" | "ComboBox" | "DropDownList" | "DatePicker" | "RepeatingSection" | "RichText" | "PlainText" | "Group";
```

Property Value  
[Word.ContentControlType](/en-us/javascript/api/word/word.contentcontroltype) | "Unknown" | "RichTextInline" | "RichTextParagraphs" | "RichTextTableCell" | "RichTextTableRow" | "RichTextTable" | "PlainTextInline" | "PlainTextParagraph" | "Picture" | "BuildingBlockGallery" | "CheckBox" | "ComboBox" | "DropDownList" | "DatePicker" | "RepeatingSection" | "RichText" | "PlainText" | "Group"

Remarks  
[API set: WordApi 1.1]

---

### xmlMapping

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns an `XmlMapping` object that represents the mapping of the content control to XML data in the data store of the document.

```typescript
xmlMapping?: Word.Interfaces.XmlMappingData;
```

Property Value  
[Word.Interfaces.XmlMappingData](/en-us/javascript/api/word/word.interfaces.xmlmappingdata)

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)]