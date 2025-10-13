# Word.Interfaces.ContentControlUpdateData interface

Package: [word](/en-us/javascript/api/word)

An interface for updating data on the ContentControl object, for use in contentControl.set({ ... }).

## Properties

- appearance  
  Specifies the appearance of the content control. The value can be 'BoundingBox', 'Tags', or 'Hidden'.

- buildingBlockGalleryContentControl  
  Gets the building block gallery-related data if the content control's [Word.ContentControlType](/en-us/javascript/api/word/word.contentcontroltype) is BuildingBlockGallery. It's null otherwise.

- cannotDelete  
  Specifies a value that indicates whether the user can delete the content control. Mutually exclusive with removeWhenEdited.

- cannotEdit  
  Specifies a value that indicates whether the user can edit the contents of the content control.

- checkboxContentControl  
  Gets the data of the content control when its type is CheckBox. It's null otherwise.

- color  
  Specifies the color of the content control. Color is specified in '#RRGGBB' format or by using the color name.

- datePickerContentControl  
  Gets the date picker-related data if the content control's [Word.ContentControlType](/en-us/javascript/api/word/word.contentcontroltype) is DatePicker. It's null otherwise.

- font  
  Gets the text format of the content control. Use this to get and set font name, size, color, and other properties.

- groupContentControl  
  Gets the group-related data if the content control's [Word.ContentControlType](/en-us/javascript/api/word/word.contentcontroltype) is Group. It's null otherwise.

- pictureContentControl  
  Gets the picture-related data if the content control's [Word.ContentControlType](/en-us/javascript/api/word/word.contentcontroltype) is Picture. It's null otherwise.

- placeholderText  
  Specifies the placeholder text of the content control. Dimmed text will be displayed when the content control is empty.

- removeWhenEdited  
  Specifies a value that indicates whether the content control is removed after it is edited. Mutually exclusive with cannotDelete.

- repeatingSectionContentControl  
  Gets the repeating section-related data if the content control's [Word.ContentControlType](/en-us/javascript/api/word/word.contentcontroltype) is RepeatingSection. It's null otherwise.

- style  
  Specifies the style name for the content control. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.

- styleBuiltIn  
  Specifies the built-in style name for the content control. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.

- tag  
  Specifies a tag to identify a content control.

- title  
  Specifies the title for a content control.

- xmlMapping  
  Returns an XmlMapping object that represents the mapping of the content control to XML data in the data store of the document.

## Property Details

### appearance

Specifies the appearance of the content control. The value can be 'BoundingBox', 'Tags', or 'Hidden'.

```typescript
appearance?: Word.ContentControlAppearance | "BoundingBox" | "Tags" | "Hidden";
```

- Property Value: [Word.ContentControlAppearance](/en-us/javascript/api/word/word.contentcontrolappearance) | "BoundingBox" | "Tags" | "Hidden"  
- Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### buildingBlockGalleryContentControl

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the building block gallery-related data if the content control's [Word.ContentControlType](/en-us/javascript/api/word/word.contentcontroltype) is `BuildingBlockGallery`. It's `null` otherwise.

```typescript
buildingBlockGalleryContentControl?: Word.Interfaces.BuildingBlockGalleryContentControlUpdateData;
```

- Property Value: [Word.Interfaces.BuildingBlockGalleryContentControlUpdateData](/en-us/javascript/api/word/word.interfaces.buildingblockgallerycontentcontrolupdatedata)  
- Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### cannotDelete

Specifies a value that indicates whether the user can delete the content control. Mutually exclusive with removeWhenEdited.

```typescript
cannotDelete?: boolean;
```

- Property Value: boolean  
- Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### cannotEdit

Specifies a value that indicates whether the user can edit the contents of the content control.

```typescript
cannotEdit?: boolean;
```

- Property Value: boolean  
- Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### checkboxContentControl

Gets the data of the content control when its type is `CheckBox`. It's `null` otherwise.

```typescript
checkboxContentControl?: Word.Interfaces.CheckboxContentControlUpdateData;
```

- Property Value: [Word.Interfaces.CheckboxContentControlUpdateData](/en-us/javascript/api/word/word.interfaces.checkboxcontentcontrolupdatedata)  
- Remarks: [API set: WordApi 1.7](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### color

Specifies the color of the content control. Color is specified in '#RRGGBB' format or by using the color name.

```typescript
color?: string;
```

- Property Value: string  
- Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### datePickerContentControl

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the date picker-related data if the content control's [Word.ContentControlType](/en-us/javascript/api/word/word.contentcontroltype) is `DatePicker`. It's `null` otherwise.

```typescript
datePickerContentControl?: Word.Interfaces.DatePickerContentControlUpdateData;
```

- Property Value: [Word.Interfaces.DatePickerContentControlUpdateData](/en-us/javascript/api/word/word.interfaces.datepickercontentcontrolupdatedata)  
- Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### font

Gets the text format of the content control. Use this to get and set font name, size, color, and other properties.

```typescript
font?: Word.Interfaces.FontUpdateData;
```

- Property Value: [Word.Interfaces.FontUpdateData](/en-us/javascript/api/word/word.interfaces.fontupdatedata)  
- Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### groupContentControl

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the group-related data if the content control's [Word.ContentControlType](/en-us/javascript/api/word/word.contentcontroltype) is `Group`. It's `null` otherwise.

```typescript
groupContentControl?: Word.Interfaces.GroupContentControlUpdateData;
```

- Property Value: [Word.Interfaces.GroupContentControlUpdateData](/en-us/javascript/api/word/word.interfaces.groupcontentcontrolupdatedata)  
- Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### pictureContentControl

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the picture-related data if the content control's [Word.ContentControlType](/en-us/javascript/api/word/word.contentcontroltype) is `Picture`. It's `null` otherwise.

```typescript
pictureContentControl?: Word.Interfaces.PictureContentControlUpdateData;
```

- Property Value: [Word.Interfaces.PictureContentControlUpdateData](/en-us/javascript/api/word/word.interfaces.picturecontentcontrolupdatedata)  
- Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### placeholderText

Specifies the placeholder text of the content control. Dimmed text will be displayed when the content control is empty.

```typescript
placeholderText?: string;
```

- Property Value: string  
- Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### removeWhenEdited

Specifies a value that indicates whether the content control is removed after it is edited. Mutually exclusive with cannotDelete.

```typescript
removeWhenEdited?: boolean;
```

- Property Value: boolean  
- Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### repeatingSectionContentControl

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the repeating section-related data if the content control's [Word.ContentControlType](/en-us/javascript/api/word/word.contentcontroltype) is `RepeatingSection`. It's `null` otherwise.

```typescript
repeatingSectionContentControl?: Word.Interfaces.RepeatingSectionContentControlUpdateData;
```

- Property Value: [Word.Interfaces.RepeatingSectionContentControlUpdateData](/en-us/javascript/api/word/word.interfaces.repeatingsectioncontentcontrolupdatedata)  
- Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### style

Specifies the style name for the content control. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.

```typescript
style?: string;
```

- Property Value: string  
- Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### styleBuiltIn

Specifies the built-in style name for the content control. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.

```typescript
styleBuiltIn?: Word.BuiltInStyleName | "Other" | "Normal" | "Heading1" | "Heading2" | "Heading3" | "Heading4" | "Heading5" | "Heading6" | "Heading7" | "Heading8" | "Heading9" | "Toc1" | "Toc2" | "Toc3" | "Toc4" | "Toc5" | "Toc6" | "Toc7" | "Toc8" | "Toc9" | "FootnoteText" | "Header" | "Footer" | "Caption" | "FootnoteReference" | "EndnoteReference" | "EndnoteText" | "Title" | "Subtitle" | "Hyperlink" | "Strong" | "Emphasis" | "NoSpacing" | "ListParagraph" | "Quote" | "IntenseQuote" | "SubtleEmphasis" | "IntenseEmphasis" | "SubtleReference" | "IntenseReference" | "BookTitle" | "Bibliography" | "TocHeading" | "TableGrid" | "PlainTable1" | "PlainTable2" | "PlainTable3" | "PlainTable4" | "PlainTable5" | "TableGridLight" | "GridTable1Light" | "GridTable1Light_Accent1" | "GridTable1Light_Accent2" | "GridTable1Light_Accent3" | "GridTable1Light_Accent4" | "GridTable1Light_Accent5" | "GridTable1Light_Accent6" | "GridTable2" | "GridTable2_Accent1" | "GridTable2_Accent2" | "GridTable2_Accent3" | "GridTable2_Accent4" | "GridTable2_Accent5" | "GridTable2_Accent6" | "GridTable3" | "GridTable3_Accent1" | "GridTable3_Accent2" | "GridTable3_Accent3" | "GridTable3_Accent4" | "GridTable3_Accent5" | "GridTable3_Accent6" | "GridTable4" | "GridTable4_Accent1" | "GridTable4_Accent2" | "GridTable4_Accent3" | "GridTable4_Accent4" | "GridTable4_Accent5" | "GridTable4_Accent6" | "GridTable5Dark" | "GridTable5Dark_Accent1" | "GridTable5Dark_Accent2" | "GridTable5Dark_Accent3" | "GridTable5Dark_Accent4" | "GridTable5Dark_Accent5" | "GridTable5Dark_Accent6" | "GridTable6Colorful" | "GridTable6Colorful_Accent1" | "GridTable6Colorful_Accent2" | "GridTable6Colorful_Accent3" | "GridTable6Colorful_Accent4" | "GridTable6Colorful_Accent5" | "GridTable6Colorful_Accent6" | "GridTable7Colorful" | "GridTable7Colorful_Accent1" | "GridTable7Colorful_Accent2" | "GridTable7Colorful_Accent3" | "GridTable7Colorful_Accent4" | "GridTable7Colorful_Accent5" | "GridTable7Colorful_Accent6" | "ListTable1Light" | "ListTable1Light_Accent1" | "ListTable1Light_Accent2" | "ListTable1Light_Accent3" | "ListTable1Light_Accent4" | "ListTable1Light_Accent5" | "ListTable1Light_Accent6" | "ListTable2" | "ListTable2_Accent1" | "ListTable2_Accent2" | "ListTable2_Accent3" | "ListTable2_Accent4" | "ListTable2_Accent5" | "ListTable2_Accent6" | "ListTable3" | "ListTable3_Accent1" | "ListTable3_Accent2" | "ListTable3_Accent3" | "ListTable3_Accent4" | "ListTable3_Accent5" | "ListTable3_Accent6" | "ListTable4" | "ListTable4_Accent1" | "ListTable4_Accent2" | "ListTable4_Accent3" | "ListTable4_Accent4" | "ListTable4_Accent5" | "ListTable4_Accent6" | "ListTable5Dark" | "ListTable5Dark_Accent1" | "ListTable5Dark_Accent2" | "ListTable5Dark_Accent3" | "ListTable5Dark_Accent4" | "ListTable5Dark_Accent5" | "ListTable5Dark_Accent6" | "ListTable6Colorful" | "ListTable6Colorful_Accent1" | "ListTable6Colorful_Accent2" | "ListTable6Colorful_Accent3" | "ListTable6Colorful_Accent4" | "ListTable6Colorful_Accent5" | "ListTable6Colorful_Accent6" | "ListTable7Colorful" | "ListTable7Colorful_Accent1" | "ListTable7Colorful_Accent2" | "ListTable7Colorful_Accent3" | "ListTable7Colorful_Accent4" | "ListTable7Colorful_Accent5" | "ListTable7Colorful_Accent6";
```

- Property Value: [Word.BuiltInStyleName](/en-us/javascript/api/word/word.builtinstylename) | "Other" | "Normal" | "Heading1" | "Heading2" | "Heading3" | "Heading4" | "Heading5" | "Heading6" | "Heading7" | "Heading8" | "Heading9" | "Toc1" | "Toc2" | "Toc3" | "Toc4" | "Toc5" | "Toc6" | "Toc7" | "Toc8" | "Toc9" | "FootnoteText" | "Header" | "Footer" | "Caption" | "FootnoteReference" | "EndnoteReference" | "EndnoteText" | "Title" | "Subtitle" | "Hyperlink" | "Strong" | "Emphasis" | "NoSpacing" | "ListParagraph" | "Quote" | "IntenseQuote" | "SubtleEmphasis" | "IntenseEmphasis" | "SubtleReference" | "IntenseReference" | "BookTitle" | "Bibliography" | "TocHeading" | "TableGrid" | "PlainTable1" | "PlainTable2" | "PlainTable3" | "PlainTable4" | "PlainTable5" | "TableGridLight" | "GridTable1Light" | "GridTable1Light_Accent1" | "GridTable1Light_Accent2" | "GridTable1Light_Accent3" | "GridTable1Light_Accent4" | "GridTable1Light_Accent5" | "GridTable1Light_Accent6" | "GridTable2" | "GridTable2_Accent1" | "GridTable2_Accent2" | "GridTable2_Accent3" | "GridTable2_Accent4" | "GridTable2_Accent5" | "GridTable2_Accent6" | "GridTable3" | "GridTable3_Accent1" | "GridTable3_Accent2" | "GridTable3_Accent3" | "GridTable3_Accent4" | "GridTable3_Accent5" | "GridTable3_Accent6" | "GridTable4" | "GridTable4_Accent1" | "GridTable4_Accent2" | "GridTable4_Accent3" | "GridTable4_Accent4" | "GridTable4_Accent5" | "GridTable4_Accent6" | "GridTable5Dark" | "GridTable5Dark_Accent1" | "GridTable5Dark_Accent2" | "GridTable5Dark_Accent3" | "GridTable5Dark_Accent4" | "GridTable5Dark_Accent5" | "GridTable5Dark_Accent6" | "GridTable6Colorful" | "GridTable6Colorful_Accent1" | "GridTable6Colorful_Accent2" | "GridTable6Colorful_Accent3" | "GridTable6Colorful_Accent4" | "GridTable6Colorful_Accent5" | "GridTable6Colorful_Accent6" | "GridTable7Colorful" | "GridTable7Colorful_Accent1" | "GridTable7Colorful_Accent2" | "GridTable7Colorful_Accent3" | "GridTable7Colorful_Accent4" | "GridTable7Colorful_Accent5" | "GridTable7Colorful_Accent6" | "ListTable1Light" | "ListTable1Light_Accent1" | "ListTable1Light_Accent2" | "ListTable1Light_Accent3" | "ListTable1Light_Accent4" | "ListTable1Light_Accent5" | "ListTable1Light_Accent6" | "ListTable2" | "ListTable2_Accent1" | "ListTable2_Accent2" | "ListTable2_Accent3" | "ListTable2_Accent4" | "ListTable2_Accent5" | "ListTable2_Accent6" | "ListTable3" | "ListTable3_Accent1" | "ListTable3_Accent2" | "ListTable3_Accent3" | "ListTable3_Accent4" | "ListTable3_Accent5" | "ListTable3_Accent6" | "ListTable4" | "ListTable4_Accent1" | "ListTable4_Accent2" | "ListTable4_Accent3" | "ListTable4_Accent4" | "ListTable4_Accent5" | "ListTable4_Accent6" | "ListTable5Dark" | "ListTable5Dark_Accent1" | "ListTable5Dark_Accent2" | "ListTable5Dark_Accent3" | "ListTable5Dark_Accent4" | "ListTable5Dark_Accent5" | "ListTable5Dark_Accent6" | "ListTable6Colorful" | "ListTable6Colorful_Accent1" | "ListTable6Colorful_Accent2" | "ListTable6Colorful_Accent3" | "ListTable6Colorful_Accent4" | "ListTable6Colorful_Accent5" | "ListTable6Colorful_Accent6" | "ListTable7Colorful" | "ListTable7Colorful_Accent1" | "ListTable7Colorful_Accent2" | "ListTable7Colorful_Accent3" | "ListTable7Colorful_Accent4" | "ListTable7Colorful_Accent5" | "ListTable7Colorful_Accent6"  
- Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### tag

Specifies a tag to identify a content control.

```typescript
tag?: string;
```

- Property Value: string  
- Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### title

Specifies the title for a content control.

```typescript
title?: string;
```

- Property Value: string  
- Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### xmlMapping

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns an `XmlMapping` object that represents the mapping of the content control to XML data in the data store of the document.

```typescript
xmlMapping?: Word.Interfaces.XmlMappingUpdateData;
```

- Property Value: [Word.Interfaces.XmlMappingUpdateData](/en-us/javascript/api/word/word.interfaces.xmlmappingupdatedata)  
- Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)