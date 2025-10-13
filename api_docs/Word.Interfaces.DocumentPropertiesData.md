# Word.Interfaces.DocumentPropertiesData interface

Package: [word](/en-us/javascript/api/word)

An interface describing the data returned by calling `documentProperties.toJSON()`.

## Properties

- [applicationName](#word-word-interfaces-documentpropertiesdata-applicationname-member): Gets the application name of the document.
- [author](#word-word-interfaces-documentpropertiesdata-author-member): Specifies the author of the document.
- [category](#word-word-interfaces-documentpropertiesdata-category-member): Specifies the category of the document.
- [comments](#word-word-interfaces-documentpropertiesdata-comments-member): Specifies the Comments field in the metadata of the document. These have no connection to comments by users made in the document.
- [company](#word-word-interfaces-documentpropertiesdata-company-member): Specifies the company of the document.
- [creationDate](#word-word-interfaces-documentpropertiesdata-creationdate-member): Gets the creation date of the document.
- [customProperties](#word-word-interfaces-documentpropertiesdata-customproperties-member): Gets the collection of custom properties of the document.
- [format](#word-word-interfaces-documentpropertiesdata-format-member): Specifies the format of the document.
- [keywords](#word-word-interfaces-documentpropertiesdata-keywords-member): Specifies the keywords of the document.
- [lastAuthor](#word-word-interfaces-documentpropertiesdata-lastauthor-member): Gets the last author of the document.
- [lastPrintDate](#word-word-interfaces-documentpropertiesdata-lastprintdate-member): Gets the last print date of the document.
- [lastSaveTime](#word-word-interfaces-documentpropertiesdata-lastsavetime-member): Gets the last save time of the document.
- [manager](#word-word-interfaces-documentpropertiesdata-manager-member): Specifies the manager of the document.
- [revisionNumber](#word-word-interfaces-documentpropertiesdata-revisionnumber-member): Gets the revision number of the document.
- [security](#word-word-interfaces-documentpropertiesdata-security-member): Gets security settings of the document. Some are access restrictions on the file on disk. Others are Document Protection settings. Some possible values are 0 = File on disk is read/write; 1 = Protect Document: File is encrypted and requires a password to open; 2 = Protect Document: Always Open as Read-Only; 3 = Protect Document: Both #1 and #2; 4 = File on disk is read-only; 5 = Both #1 and #4; 6 = Both #2 and #4; 7 = All of #1, #2, and #4; 8 = Protect Document: Restrict Edit to read-only; 9 = Both #1 and #8; 10 = Both #2 and #8; 11 = All of #1, #2, and #8; 12 = Both #4 and #8; 13 = All of #1, #4, and #8; 14 = All of #2, #4, and #8; 15 = All of #1, #2, #4, and #8.
- [subject](#word-word-interfaces-documentpropertiesdata-subject-member): Specifies the subject of the document.
- [template](#word-word-interfaces-documentpropertiesdata-template-member): Gets the template of the document.
- [title](#word-word-interfaces-documentpropertiesdata-title-member): Specifies the title of the document.

## Property Details

<a id="word-word-interfaces-documentpropertiesdata-applicationname-member"></a>
### applicationName

Gets the application name of the document.

```typescript
applicationName?: string;
```

Property Value: string

Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-documentpropertiesdata-author-member"></a>
### author

Specifies the author of the document.

```typescript
author?: string;
```

Property Value: string

Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-documentpropertiesdata-category-member"></a>
### category

Specifies the category of the document.

```typescript
category?: string;
```

Property Value: string

Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-documentpropertiesdata-comments-member"></a>
### comments

Specifies the Comments field in the metadata of the document. These have no connection to comments by users made in the document.

```typescript
comments?: string;
```

Property Value: string

Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-documentpropertiesdata-company-member"></a>
### company

Specifies the company of the document.

```typescript
company?: string;
```

Property Value: string

Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-documentpropertiesdata-creationdate-member"></a>
### creationDate

Gets the creation date of the document.

```typescript
creationDate?: Date;
```

Property Value: Date

Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-documentpropertiesdata-customproperties-member"></a>
### customProperties

Gets the collection of custom properties of the document.

```typescript
customProperties?: Word.Interfaces.CustomPropertyData[];
```

Property Value: [Word.Interfaces.CustomPropertyData](/en-us/javascript/api/word/word.interfaces.custompropertydata)[]

Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-documentpropertiesdata-format-member"></a>
### format

Specifies the format of the document.

```typescript
format?: string;
```

Property Value: string

Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-documentpropertiesdata-keywords-member"></a>
### keywords

Specifies the keywords of the document.

```typescript
keywords?: string;
```

Property Value: string

Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-documentpropertiesdata-lastauthor-member"></a>
### lastAuthor

Gets the last author of the document.

```typescript
lastAuthor?: string;
```

Property Value: string

Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-documentpropertiesdata-lastprintdate-member"></a>
### lastPrintDate

Gets the last print date of the document.

```typescript
lastPrintDate?: Date;
```

Property Value: Date

Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-documentpropertiesdata-lastsavetime-member"></a>
### lastSaveTime

Gets the last save time of the document.

```typescript
lastSaveTime?: Date;
```

Property Value: Date

Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-documentpropertiesdata-manager-member"></a>
### manager

Specifies the manager of the document.

```typescript
manager?: string;
```

Property Value: string

Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-documentpropertiesdata-revisionnumber-member"></a>
### revisionNumber

Gets the revision number of the document.

```typescript
revisionNumber?: string;
```

Property Value: string

Remarks: [API set: WordApi 1.3](/en-us/office/dev/add-ins/reference/overview/visio-javascript-reference-overview)

<a id="word-word-interfaces-documentpropertiesdata-security-member"></a>
### security

Gets security settings of the document. Some are access restrictions on the file on disk. Others are Document Protection settings. Some possible values are 0 = File on disk is read/write; 1 = Protect Document: File is encrypted and requires a password to open; 2 = Protect Document: Always Open as Read-Only; 3 = Protect Document: Both #1 and #2; 4 = File on disk is read-only; 5 = Both #1 and #4; 6 = Both #2 and #4; 7 = All of #1, #2, and #4; 8 = Protect Document: Restrict Edit to read-only; 9 = Both #1 and #8; 10 = Both #2 and #8; 11 = All of #1, #2, and #8; 12 = Both #4 and #8; 13 = All of #1, #4, and #8; 14 = All of #2, #4, and #8; 15 = All of #1, #2, #4, and #8.

```typescript
security?: number;
```

Property Value: number

Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-documentpropertiesdata-subject-member"></a>
### subject

Specifies the subject of the document.

```typescript
subject?: string;
```

Property Value: string

Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-documentpropertiesdata-template-member"></a>
### template

Gets the template of the document.

```typescript
template?: string;
```

Property Value: string

Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-documentpropertiesdata-title-member"></a>
### title

Specifies the title of the document.

```typescript
title?: string;
```

Property Value: string

Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)