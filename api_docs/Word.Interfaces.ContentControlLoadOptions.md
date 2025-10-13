# Word.Interfaces.ContentControlLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Represents a content control. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain contents such as images, tables, or paragraphs of formatted text. Currently, only rich text, plain text, checkbox, dropdown list, and combo box content controls are supported.

## Remarks

[API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- $all  
  Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

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

- datePickerContentControl  
  Gets the date picker-related data if the content control's [Word.ContentControlType](/en-us/javascript/api/word/word.contentcontroltype) is `DatePicker`. It's `null` otherwise.

- font  
  Gets the text format of the content control. Use this to get and set font name, size, color, and other properties.

- groupContentControl  
  Gets the group-related data if the content control's [Word.ContentControlType](/en-us/javascript/api/word/word.contentcontroltype) is `Group`. It's `null` otherwise.

- id  
  Gets an integer that represents the content control identifier.

- parentBody  
  Gets the parent body of the content control.

- parentContentControl  
  Gets the content control that contains the content control. Throws an `ItemNotFound` error if there isn't a parent content control.

- parentContentControlOrNullObject  
  Gets the content control that contains the content control. If there isn't a parent content control, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

- parentTable  
  Gets the table that contains the content control. Throws an `ItemNotFound` error if it isn't contained in a table.

- parentTableCell  
  Gets the table cell that contains the content control. Throws an `ItemNotFound` error if it isn't contained in a table cell.

- parentTableCellOrNullObject  
  Gets the table cell that contains the content control. If it isn't contained in a table cell, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

- parentTableOrNullObject  
  Gets the table that contains the content control. If it isn't contained in a table, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

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

### $all

Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

- Property value: boolean

---

### appearance

Specifies the appearance of the content control. The value can be 'BoundingBox', 'Tags', or 'Hidden'.

```typescript
appearance?: boolean;
```

- Property value: boolean

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### buildingBlockGalleryContentControl

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the building block gallery-related data if the content control's [Word.ContentControlType](/en-us/javascript/api/word/word.contentcontroltype) is `BuildingBlockGallery`. It's `null` otherwise.

```typescript
buildingBlockGalleryContentControl?: Word.Interfaces.BuildingBlockGalleryContentControlLoadOptions;
```

- Property value: [Word.Interfaces.BuildingBlockGalleryContentControlLoadOptions](/en-us/javascript/api/word/word.interfaces.buildingblockgallerycontentcontrolloadoptions)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### cannotDelete

Specifies a value that indicates whether the user can delete the content control. Mutually exclusive with removeWhenEdited.

```typescript
cannotDelete?: boolean;
```

- Property value: boolean

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### cannotEdit

Specifies a value that indicates whether the user can edit the contents of the content control.

```typescript
cannotEdit?: boolean;
```

- Property value: boolean

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### checkboxContentControl

Gets the data of the content control when its type is `CheckBox`. It's `null` otherwise.

```typescript
checkboxContentControl?: Word.Interfaces.CheckboxContentControlLoadOptions;
```

- Property value: [Word.Interfaces.CheckboxContentControlLoadOptions](/en-us/javascript/api/word/word.interfaces.checkboxcontentcontrolloadoptions)

Remarks: [API set: WordApi 1.7](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### color

Specifies the color of the content control. Color is specified in '#RRGGBB' format or by using the color name.

```typescript
color?: boolean;
```

- Property value: boolean

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### datePickerContentControl

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the date picker-related data if the content control's [Word.ContentControlType](/en-us/javascript/api/word/word.contentcontroltype) is `DatePicker`. It's `null` otherwise.

```typescript
datePickerContentControl?: Word.Interfaces.DatePickerContentControlLoadOptions;
```

- Property value: [Word.Interfaces.DatePickerContentControlLoadOptions](/en-us/javascript/api/word/word.interfaces.datepickercontentcontrolloadoptions)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### font

Gets the text format of the content control. Use this to get and set font name, size, color, and other properties.

```typescript
font?: Word.Interfaces.FontLoadOptions;
```

- Property value: [Word.Interfaces.FontLoadOptions](/en-us/javascript/api/word/word.interfaces.fontloadoptions)

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### groupContentControl

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the group-related data if the content control's [Word.ContentControlType](/en-us/javascript/api/word/word.contentcontroltype) is `Group`. It's `null` otherwise.

```typescript
groupContentControl?: Word.Interfaces.GroupContentControlLoadOptions;
```

- Property value: [Word.Interfaces.GroupContentControlLoadOptions](/en-us/javascript/api/word/word.interfaces.groupcontentcontrolloadoptions)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### id

Gets an integer that represents the content control identifier.

```typescript
id?: boolean;
```

- Property value: boolean

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### parentBody

Gets the parent body of the content control.

```typescript
parentBody?: Word.Interfaces.BodyLoadOptions;
```

- Property value: [Word.Interfaces.BodyLoadOptions](/en-us/javascript/api/word/word.interfaces.bodyloadoptions)

Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### parentContentControl

Gets the content control that contains the content control. Throws an `ItemNotFound` error if there isn't a parent content control.

```typescript
parentContentControl?: Word.Interfaces.ContentControlLoadOptions;
```

- Property value: [Word.Interfaces.ContentControlLoadOptions](/en-us/javascript/api/word/word.interfaces.contentcontrolloadoptions)

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### parentContentControlOrNullObject

Gets the content control that contains the content control. If there isn't a parent content control, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
parentContentControlOrNullObject?: Word.Interfaces.ContentControlLoadOptions;
```

- Property value: [Word.Interfaces.ContentControlLoadOptions](/en-us/javascript/api/word/word.interfaces.contentcontrolloadoptions)

Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### parentTable

Gets the table that contains the content control. Throws an `ItemNotFound` error if it isn't contained in a table.

```typescript
parentTable?: Word.Interfaces.TableLoadOptions;
```

- Property value: [Word.Interfaces.TableLoadOptions](/en-us/javascript/api/word/word.interfaces.tableloadoptions)

Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### parentTableCell

Gets the table cell that contains the content control. Throws an `ItemNotFound` error if it isn't contained in a table cell.

```typescript
parentTableCell?: Word.Interfaces.TableCellLoadOptions;
```

- Property value: [Word.Interfaces.TableCellLoadOptions](/en-us/javascript/api/word/word.interfaces.tablecellloadoptions)

Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### parentTableCellOrNullObject

Gets the table cell that contains the content control. If it isn't contained in a table cell, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
parentTableCellOrNullObject?: Word.Interfaces.TableCellLoadOptions;
```

- Property value: [Word.Interfaces.TableCellLoadOptions](/en-us/javascript/api/word/word.interfaces.tablecellloadoptions)

Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### parentTableOrNullObject

Gets the table that contains the content control. If it isn't contained in a table, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
parentTableOrNullObject?: Word.Interfaces.TableLoadOptions;
```

- Property value: [Word.Interfaces.TableLoadOptions](/en-us/javascript/api/word/word.interfaces.tableloadoptions)

Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### pictureContentControl

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the picture-related data if the content control's [Word.ContentControlType](/en-us/javascript/api/word/word.contentcontroltype) is `Picture`. It's `null` otherwise.

```typescript
pictureContentControl?: Word.Interfaces.PictureContentControlLoadOptions;
```

- Property value: [Word.Interfaces.PictureContentControlLoadOptions](/en-us/javascript/api/word/word.interfaces.picturecontentcontrolloadoptions)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### placeholderText

Specifies the placeholder text of the content control. Dimmed text will be displayed when the content control is empty.

```typescript
placeholderText?: boolean;
```

- Property value: boolean

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### removeWhenEdited

Specifies a value that indicates whether the content control is removed after it is edited. Mutually exclusive with cannotDelete.

```typescript
removeWhenEdited?: boolean;
```

- Property value: boolean

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### repeatingSectionContentControl

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the repeating section-related data if the content control's [Word.ContentControlType](/en-us/javascript/api/word/word.contentcontroltype) is `RepeatingSection`. It's `null` otherwise.

```typescript
repeatingSectionContentControl?: Word.Interfaces.RepeatingSectionContentControlLoadOptions;
```

- Property value: [Word.Interfaces.RepeatingSectionContentControlLoadOptions](/en-us/javascript/api/word/word.interfaces.repeatingsectioncontentcontrolloadoptions)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### style

Specifies the style name for the content control. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.

```typescript
style?: boolean;
```

- Property value: boolean

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### styleBuiltIn

Specifies the built-in style name for the content control. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.

```typescript
styleBuiltIn?: boolean;
```

- Property value: boolean

Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### subtype

Gets the content control subtype. The subtype can be 'RichTextInline', 'RichTextParagraphs', 'RichTextTableCell', 'RichTextTableRow' and 'RichTextTable' for rich text content controls, or 'PlainTextInline' and 'PlainTextParagraph' for plain text content controls, or 'CheckBox' for checkbox content controls.

```typescript
subtype?: boolean;
```

- Property value: boolean

Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### tag

Specifies a tag to identify a content control.

```typescript
tag?: boolean;
```

- Property value: boolean

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### text

Gets the text of the content control.

```typescript
text?: boolean;
```

- Property value: boolean

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### title

Specifies the title for a content control.

```typescript
title?: boolean;
```

- Property value: boolean

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### type

Gets the content control type. Only rich text, plain text, and checkbox content controls are supported currently.

```typescript
type?: boolean;
```

- Property value: boolean

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### xmlMapping

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns an `XmlMapping` object that represents the mapping of the content control to XML data in the data store of the document.

```typescript
xmlMapping?: Word.Interfaces.XmlMappingLoadOptions;
```

- Property value: [Word.Interfaces.XmlMappingLoadOptions](/en-us/javascript/api/word/word.interfaces.xmlmappingloadoptions)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)