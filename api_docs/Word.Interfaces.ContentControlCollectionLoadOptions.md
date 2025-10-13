# Word.Interfaces.ContentControlCollectionLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Contains a collection of [Word.ContentControl](/en-us/javascript/api/word/word.contentcontrol) objects. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain contents such as images, tables, or paragraphs of formatted text. Currently, only rich text, plain text, checkbox, dropdown list, and combo box content controls are supported.

## Remarks

[ API set: WordApi 1.1 ]

## Properties

- $all  
  Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

- appearance  
  For EACH ITEM in the collection: Specifies the appearance of the content control. The value can be 'BoundingBox', 'Tags', or 'Hidden'.

- buildingBlockGalleryContentControl  
  For EACH ITEM in the collection: Gets the building block gallery-related data if the content control's [Word.ContentControlType](/en-us/javascript/api/word/word.contentcontroltype) is `BuildingBlockGallery`. It's `null` otherwise.

- cannotDelete  
  For EACH ITEM in the collection: Specifies a value that indicates whether the user can delete the content control. Mutually exclusive with removeWhenEdited.

- cannotEdit  
  For EACH ITEM in the collection: Specifies a value that indicates whether the user can edit the contents of the content control.

- checkboxContentControl  
  For EACH ITEM in the collection: Gets the data of the content control when its type is `CheckBox`. It's `null` otherwise.

- color  
  For EACH ITEM in the collection: Specifies the color of the content control. Color is specified in '#RRGGBB' format or by using the color name.

- datePickerContentControl  
  For EACH ITEM in the collection: Gets the date picker-related data if the content control's [Word.ContentControlType](/en-us/javascript/api/word/word.contentcontroltype) is `DatePicker`. It's `null` otherwise.

- font  
  For EACH ITEM in the collection: Gets the text format of the content control. Use this to get and set font name, size, color, and other properties.

- groupContentControl  
  For EACH ITEM in the collection: Gets the group-related data if the content control's [Word.ContentControlType](/en-us/javascript/api/word/word.contentcontroltype) is `Group`. It's `null` otherwise.

- id  
  For EACH ITEM in the collection: Gets an integer that represents the content control identifier.

- parentBody  
  For EACH ITEM in the collection: Gets the parent body of the content control.

- parentContentControl  
  For EACH ITEM in the collection: Gets the content control that contains the content control. Throws an `ItemNotFound` error if there isn't a parent content control.

- parentContentControlOrNullObject  
  For EACH ITEM in the collection: Gets the content control that contains the content control. If there isn't a parent content control, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

- parentTable  
  For EACH ITEM in the collection: Gets the table that contains the content control. Throws an `ItemNotFound` error if it isn't contained in a table.

- parentTableCell  
  For EACH ITEM in the collection: Gets the table cell that contains the content control. Throws an `ItemNotFound` error if it isn't contained in a table cell.

- parentTableCellOrNullObject  
  For EACH ITEM in the collection: Gets the table cell that contains the content control. If it isn't contained in a table cell, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

- parentTableOrNullObject  
  For EACH ITEM in the collection: Gets the table that contains the content control. If it isn't contained in a table, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

- pictureContentControl  
  For EACH ITEM in the collection: Gets the picture-related data if the content control's [Word.ContentControlType](/en-us/javascript/api/word/word.contentcontroltype) is `Picture`. It's `null` otherwise.

- placeholderText  
  For EACH ITEM in the collection: Specifies the placeholder text of the content control. Dimmed text will be displayed when the content control is empty.

- removeWhenEdited  
  For EACH ITEM in the collection: Specifies a value that indicates whether the content control is removed after it is edited. Mutually exclusive with cannotDelete.

- repeatingSectionContentControl  
  For EACH ITEM in the collection: Gets the repeating section-related data if the content control's [Word.ContentControlType](/en-us/javascript/api/word/word.contentcontroltype) is `RepeatingSection`. It's `null` otherwise.

- style  
  For EACH ITEM in the collection: Specifies the style name for the content control. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.

- styleBuiltIn  
  For EACH ITEM in the collection: Specifies the built-in style name for the content control. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.

- subtype  
  For EACH ITEM in the collection: Gets the content control subtype. The subtype can be 'RichTextInline', 'RichTextParagraphs', 'RichTextTableCell', 'RichTextTableRow' and 'RichTextTable' for rich text content controls, or 'PlainTextInline' and 'PlainTextParagraph' for plain text content controls, or 'CheckBox' for checkbox content controls.

- tag  
  For EACH ITEM in the collection: Specifies a tag to identify a content control.

- text  
  For EACH ITEM in the collection: Gets the text of the content control.

- title  
  For EACH ITEM in the collection: Specifies the title for a content control.

- type  
  For EACH ITEM in the collection: Gets the content control type. Only rich text, plain text, and checkbox content controls are supported currently.

- xmlMapping  
  For EACH ITEM in the collection: Returns an `XmlMapping` object that represents the mapping of the content control to XML data in the data store of the document.

## Property Details

### $all

Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

Property Value  
boolean

### appearance

For EACH ITEM in the collection: Specifies the appearance of the content control. The value can be 'BoundingBox', 'Tags', or 'Hidden'.

```typescript
appearance?: boolean;
```

Property Value  
boolean

Remarks  
[ API set: WordApi 1.1 ]

### buildingBlockGalleryContentControl

Note

This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Gets the building block gallery-related data if the content control's [Word.ContentControlType](/en-us/javascript/api/word/word.contentcontroltype) is `BuildingBlockGallery`. It's `null` otherwise.

```typescript
buildingBlockGalleryContentControl?: Word.Interfaces.BuildingBlockGalleryContentControlLoadOptions;
```

Property Value  
[Word.Interfaces.BuildingBlockGalleryContentControlLoadOptions](/en-us/javascript/api/word/word.interfaces.buildingblockgallerycontentcontrolloadoptions)

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

### cannotDelete

For EACH ITEM in the collection: Specifies a value that indicates whether the user can delete the content control. Mutually exclusive with removeWhenEdited.

```typescript
cannotDelete?: boolean;
```

Property Value  
boolean

Remarks  
[ API set: WordApi 1.1 ]

### cannotEdit

For EACH ITEM in the collection: Specifies a value that indicates whether the user can edit the contents of the content control.

```typescript
cannotEdit?: boolean;
```

Property Value  
boolean

Remarks  
[ API set: WordApi 1.1 ]

### checkboxContentControl

For EACH ITEM in the collection: Gets the data of the content control when its type is `CheckBox`. It's `null` otherwise.

```typescript
checkboxContentControl?: Word.Interfaces.CheckboxContentControlLoadOptions;
```

Property Value  
[Word.Interfaces.CheckboxContentControlLoadOptions](/en-us/javascript/api/word/word.interfaces.checkboxcontentcontrolloadoptions)

Remarks  
[ API set: WordApi 1.7 ]

### color

For EACH ITEM in the collection: Specifies the color of the content control. Color is specified in '#RRGGBB' format or by using the color name.

```typescript
color?: boolean;
```

Property Value  
boolean

Remarks  
[ API set: WordApi 1.1 ]

### datePickerContentControl

Note

This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Gets the date picker-related data if the content control's [Word.ContentControlType](/en-us/javascript/api/word/word.contentcontroltype) is `DatePicker`. It's `null` otherwise.

```typescript
datePickerContentControl?: Word.Interfaces.DatePickerContentControlLoadOptions;
```

Property Value  
[Word.Interfaces.DatePickerContentControlLoadOptions](/en-us/javascript/api/word/word.interfaces.datepickercontentcontrolloadoptions)

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

### font

For EACH ITEM in the collection: Gets the text format of the content control. Use this to get and set font name, size, color, and other properties.

```typescript
font?: Word.Interfaces.FontLoadOptions;
```

Property Value  
[Word.Interfaces.FontLoadOptions](/en-us/javascript/api/word/word.interfaces.fontloadoptions)

Remarks  
[ API set: WordApi 1.1 ]

### groupContentControl

Note

This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Gets the group-related data if the content control's [Word.ContentControlType](/en-us/javascript/api/word/word.contentcontroltype) is `Group`. It's `null` otherwise.

```typescript
groupContentControl?: Word.Interfaces.GroupContentControlLoadOptions;
```

Property Value  
[Word.Interfaces.GroupContentControlLoadOptions](/en-us/javascript/api/word/word.interfaces.groupcontentcontrolloadoptions)

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

### id

For EACH ITEM in the collection: Gets an integer that represents the content control identifier.

```typescript
id?: boolean;
```

Property Value  
boolean

Remarks  
[ API set: WordApi 1.1 ]

### parentBody

For EACH ITEM in the collection: Gets the parent body of the content control.

```typescript
parentBody?: Word.Interfaces.BodyLoadOptions;
```

Property Value  
[Word.Interfaces.BodyLoadOptions](/en-us/javascript/api/word/word.interfaces.bodyloadoptions)

Remarks  
[ API set: WordApi 1.3 ]

### parentContentControl

For EACH ITEM in the collection: Gets the content control that contains the content control. Throws an `ItemNotFound` error if there isn't a parent content control.

```typescript
parentContentControl?: Word.Interfaces.ContentControlLoadOptions;
```

Property Value  
[Word.Interfaces.ContentControlLoadOptions](/en-us/javascript/api/word/word.interfaces.contentcontrolloadoptions)

Remarks  
[ API set: WordApi 1.1 ]

### parentContentControlOrNullObject

For EACH ITEM in the collection: Gets the content control that contains the content control. If there isn't a parent content control, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
parentContentControlOrNullObject?: Word.Interfaces.ContentControlLoadOptions;
```

Property Value  
[Word.Interfaces.ContentControlLoadOptions](/en-us/javascript/api/word/word.interfaces.contentcontrolloadoptions)

Remarks  
[ API set: WordApi 1.3 ]

### parentTable

For EACH ITEM in the collection: Gets the table that contains the content control. Throws an `ItemNotFound` error if it isn't contained in a table.

```typescript
parentTable?: Word.Interfaces.TableLoadOptions;
```

Property Value  
[Word.Interfaces.TableLoadOptions](/en-us/javascript/api/word/word.interfaces.tableloadoptions)

Remarks  
[ API set: WordApi 1.3 ]

### parentTableCell

For EACH ITEM in the collection: Gets the table cell that contains the content control. Throws an `ItemNotFound` error if it isn't contained in a table cell.

```typescript
parentTableCell?: Word.Interfaces.TableCellLoadOptions;
```

Property Value  
[Word.Interfaces.TableCellLoadOptions](/en-us/javascript/api/word/word.interfaces.tablecellloadoptions)

Remarks  
[ API set: WordApi 1.3 ]

### parentTableCellOrNullObject

For EACH ITEM in the collection: Gets the table cell that contains the content control. If it isn't contained in a table cell, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
parentTableCellOrNullObject?: Word.Interfaces.TableCellLoadOptions;
```

Property Value  
[Word.Interfaces.TableCellLoadOptions](/en-us/javascript/api/word/word.interfaces.tablecellloadoptions)

Remarks  
[ API set: WordApi 1.3 ]

### parentTableOrNullObject

For EACH ITEM in the collection: Gets the table that contains the content control. If it isn't contained in a table, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
parentTableOrNullObject?: Word.Interfaces.TableLoadOptions;
```

Property Value  
[Word.Interfaces.TableLoadOptions](/en-us/javascript/api/word/word.interfaces.tableloadoptions)

Remarks  
[ API set: WordApi 1.3 ]

### pictureContentControl

Note

This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Gets the picture-related data if the content control's [Word.ContentControlType](/en-us/javascript/api/word/word.contentcontroltype) is `Picture`. It's `null` otherwise.

```typescript
pictureContentControl?: Word.Interfaces.PictureContentControlLoadOptions;
```

Property Value  
[Word.Interfaces.PictureContentControlLoadOptions](/en-us/javascript/api/word/word.interfaces.picturecontentcontrolloadoptions)

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

### placeholderText

For EACH ITEM in the collection: Specifies the placeholder text of the content control. Dimmed text will be displayed when the content control is empty.

```typescript
placeholderText?: boolean;
```

Property Value  
boolean

Remarks  
[ API set: WordApi 1.1 ]

### removeWhenEdited

For EACH ITEM in the collection: Specifies a value that indicates whether the content control is removed after it is edited. Mutually exclusive with cannotDelete.

```typescript
removeWhenEdited?: boolean;
```

Property Value  
boolean

Remarks  
[ API set: WordApi 1.1 ]

### repeatingSectionContentControl

Note

This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Gets the repeating section-related data if the content control's [Word.ContentControlType](/en-us/javascript/api/word/word.contentcontroltype) is `RepeatingSection`. It's `null` otherwise.

```typescript
repeatingSectionContentControl?: Word.Interfaces.RepeatingSectionContentControlLoadOptions;
```

Property Value  
[Word.Interfaces.RepeatingSectionContentControlLoadOptions](/en-us/javascript/api/word/word.interfaces.repeatingsectioncontentcontrolloadoptions)

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

### style

For EACH ITEM in the collection: Specifies the style name for the content control. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.

```typescript
style?: boolean;
```

Property Value  
boolean

Remarks  
[ API set: WordApi 1.1 ]

### styleBuiltIn

For EACH ITEM in the collection: Specifies the built-in style name for the content control. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.

```typescript
styleBuiltIn?: boolean;
```

Property Value  
boolean

Remarks  
[ API set: WordApi 1.3 ]

### subtype

For EACH ITEM in the collection: Gets the content control subtype. The subtype can be 'RichTextInline', 'RichTextParagraphs', 'RichTextTableCell', 'RichTextTableRow' and 'RichTextTable' for rich text content controls, or 'PlainTextInline' and 'PlainTextParagraph' for plain text content controls, or 'CheckBox' for checkbox content controls.

```typescript
subtype?: boolean;
```

Property Value  
boolean

Remarks  
[ API set: WordApi 1.3 ]

### tag

For EACH ITEM in the collection: Specifies a tag to identify a content control.

```typescript
tag?: boolean;
```

Property Value  
boolean

Remarks  
[ API set: WordApi 1.1 ]

### text

For EACH ITEM in the collection: Gets the text of the content control.

```typescript
text?: boolean;
```

Property Value  
boolean

Remarks  
[ API set: WordApi 1.1 ]

### title

For EACH ITEM in the collection: Specifies the title for a content control.

```typescript
title?: boolean;
```

Property Value  
boolean

Remarks  
[ API set: WordApi 1.1 ]

### type

For EACH ITEM in the collection: Gets the content control type. Only rich text, plain text, and checkbox content controls are supported currently.

```typescript
type?: boolean;
```

Property Value  
boolean

Remarks  
[ API set: WordApi 1.1 ]

### xmlMapping

Note

This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Returns an `XmlMapping` object that represents the mapping of the content control to XML data in the data store of the document.

```typescript
xmlMapping?: Word.Interfaces.XmlMappingLoadOptions;
```

Property Value  
[Word.Interfaces.XmlMappingLoadOptions](/en-us/javascript/api/word/word.interfaces.xmlmappingloadoptions)

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]