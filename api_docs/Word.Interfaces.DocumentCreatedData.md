# Word.Interfaces.DocumentCreatedData interface

Package: https://learn.microsoft.com/en-us/javascript/api/word

An interface describing the data returned by calling `documentCreated.toJSON()`.

## Properties

- body  
  Gets the body object of the document. The body is the text that excludes headers, footers, footnotes, textboxes, etc.
- contentControls  
  Gets the collection of content control objects in the document. This includes content controls in the body of the document, headers, footers, textboxes, etc.
- customXmlParts  
  Gets the custom XML parts in the document.
- properties  
  Gets the properties of the document.
- saved  
  Indicates whether the changes in the document have been saved. A value of true indicates that the document hasn't changed since it was saved.
- sections  
  Gets the collection of section objects in the document.
- settings  
  Gets the add-in's settings in the document.

## Property Details

### body

Gets the body object of the document. The body is the text that excludes headers, footers, footnotes, textboxes, etc.

```typescript
body?: Word.Interfaces.BodyData;
```

#### Property Value
- Word.Interfaces.BodyData: https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.bodydata

#### Remarks
- [API set: WordApiHiddenDocument 1.3](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### contentControls

Gets the collection of content control objects in the document. This includes content controls in the body of the document, headers, footers, textboxes, etc.

```typescript
contentControls?: Word.Interfaces.ContentControlData[];
```

#### Property Value
- Word.Interfaces.ContentControlData[]: https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.contentcontroldata

#### Remarks
- [API set: WordApiHiddenDocument 1.3](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### customXmlParts

Gets the custom XML parts in the document.

```typescript
customXmlParts?: Word.Interfaces.CustomXmlPartData[];
```

#### Property Value
- Word.Interfaces.CustomXmlPartData[]: https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.customxmlpartdata

#### Remarks
- [API set: WordApiHiddenDocument 1.4](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### properties

Gets the properties of the document.

```typescript
properties?: Word.Interfaces.DocumentPropertiesData;
```

#### Property Value
- Word.Interfaces.DocumentPropertiesData: https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.documentpropertiesdata

#### Remarks
- [API set: WordApiHiddenDocument 1.3](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### saved

Indicates whether the changes in the document have been saved. A value of true indicates that the document hasn't changed since it was saved.

```typescript
saved?: boolean;
```

#### Property Value
- boolean

#### Remarks
- [API set: WordApiHiddenDocument 1.3](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### sections

Gets the collection of section objects in the document.

```typescript
sections?: Word.Interfaces.SectionData[];
```

#### Property Value
- Word.Interfaces.SectionData[]: https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.sectiondata

#### Remarks
- [API set: WordApiHiddenDocument 1.3](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### settings

Gets the add-in's settings in the document.

```typescript
settings?: Word.Interfaces.SettingData[];
```

#### Property Value
- Word.Interfaces.SettingData[]: https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.settingdata

#### Remarks
- [API set: WordApiHiddenDocument 1.4](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)