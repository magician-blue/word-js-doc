# Word.DocumentProperties class

Package: [word](/en-us/javascript/api/word)

Represents document properties.

Extends: [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks

[ [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

### Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/30-properties/get-built-in-properties.yaml

await Word.run(async (context) => {
    const builtInProperties: Word.DocumentProperties = context.document.properties;
    builtInProperties.load("*"); // Let's get all!

    await context.sync();
    console.log(JSON.stringify(builtInProperties, null, 4));
});
```

## Properties

- applicationName: Gets the application name of the document.
- author: Specifies the author of the document.
- category: Specifies the category of the document.
- comments: Specifies the Comments field in the metadata of the document. These have no connection to comments by users made in the document.
- company: Specifies the company of the document.
- context: The request context associated with the object. This connects the add-in's process to the Office host application's process.
- creationDate: Gets the creation date of the document.
- customProperties: Gets the collection of custom properties of the document.
- format: Specifies the format of the document.
- keywords: Specifies the keywords of the document.
- lastAuthor: Gets the last author of the document.
- lastPrintDate: Gets the last print date of the document.
- lastSaveTime: Gets the last save time of the document.
- manager: Specifies the manager of the document.
- revisionNumber: Gets the revision number of the document.
- security: Gets security settings of the document. Some are access restrictions on the file on disk. Others are Document Protection settings. Some possible values are 0 = File on disk is read/write; 1 = Protect Document: File is encrypted and requires a password to open; 2 = Protect Document: Always Open as Read-Only; 3 = Protect Document: Both #1 and #2; 4 = File on disk is read-only; 5 = Both #1 and #4; 6 = Both #2 and #4; 7 = All of #1, #2, and #4; 8 = Protect Document: Restrict Edit to read-only; 9 = Both #1 and #8; 10 = Both #2 and #8; 11 = All of #1, #2, and #8; 12 = Both #4 and #8; 13 = All of #1, #4, and #8; 14 = All of #2, #4, and #8; 15 = All of #1, #2, #4, and #8.
- subject: Specifies the subject of the document.
- template: Gets the template of the document.
- title: Specifies the title of the document.

## Methods

- load(options): Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- load(propertyNames): Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- load(propertyNamesAndPaths): Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- set(properties, options): Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
- set(properties): Sets multiple properties on the object at the same time, based on an existing loaded object.
- toJSON(): Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.DocumentProperties` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.DocumentPropertiesData`) that contains shallow copies of any loaded child properties from the original object.
- track(): Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
- untrack(): Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

## Property Details

### applicationName

Gets the application name of the document.

```typescript
readonly applicationName: string;
```

#### Property Value
- string

#### Remarks
[ [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

### author

Specifies the author of the document.

```typescript
author: string;
```

#### Property Value
- string

#### Remarks
[ [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

### category

Specifies the category of the document.

```typescript
category: string;
```

#### Property Value
- string

#### Remarks
[ [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

### comments

Specifies the Comments field in the metadata of the document. These have no connection to comments by users made in the document.

```typescript
comments: string;
```

#### Property Value
- string

#### Remarks
[ [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

### company

Specifies the company of the document.

```typescript
company: string;
```

#### Property Value
- string

#### Remarks
[ [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

### context

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

#### Property Value
- [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### creationDate

Gets the creation date of the document.

```typescript
readonly creationDate: Date;
```

#### Property Value
- Date

#### Remarks
[ [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

### customProperties

Gets the collection of custom properties of the document.

```typescript
readonly customProperties: Word.CustomPropertyCollection;
```

#### Property Value
- [Word.CustomPropertyCollection](/en-us/javascript/api/word/word.custompropertycollection)

#### Remarks
[ [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

### format

Specifies the format of the document.

```typescript
format: string;
```

#### Property Value
- string

#### Remarks
[ [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

### keywords

Specifies the keywords of the document.

```typescript
keywords: string;
```

#### Property Value
- string

#### Remarks
[ [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

### lastAuthor

Gets the last author of the document.

```typescript
readonly lastAuthor: string;
```

#### Property Value
- string

#### Remarks
[ [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

### lastPrintDate

Gets the last print date of the document.

```typescript
readonly lastPrintDate: Date;
```

#### Property Value
- Date

#### Remarks
[ [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

### lastSaveTime

Gets the last save time of the document.

```typescript
readonly lastSaveTime: Date;
```

#### Property Value
- Date

#### Remarks
[ [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

### manager

Specifies the manager of the document.

```typescript
manager: string;
```

#### Property Value
- string

#### Remarks
[ [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

### revisionNumber

Gets the revision number of the document.

```typescript
readonly revisionNumber: string;
```

#### Property Value
- string

#### Remarks
[ [API set: WordApi 1.3](/en-us/office/dev/add-ins/reference/overview/visio-javascript-reference-overview) ]

### security

Gets security settings of the document. Some are access restrictions on the file on disk. Others are Document Protection settings. Some possible values are 0 = File on disk is read/write; 1 = Protect Document: File is encrypted and requires a password to open; 2 = Protect Document: Always Open as Read-Only; 3 = Protect Document: Both #1 and #2; 4 = File on disk is read-only; 5 = Both #1 and #4; 6 = Both #2 and #4; 7 = All of #1, #2, and #4; 8 = Protect Document: Restrict Edit to read-only; 9 = Both #1 and #8; 10 = Both #2 and #8; 11 = All of #1, #2, and #8; 12 = Both #4 and #8; 13 = All of #1, #4, and #8; 14 = All of #2, #4, and #8; 15 = All of #1, #2, #4, and #8.

```typescript
readonly security: number;
```

#### Property Value
- number

#### Remarks
[ [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

### subject

Specifies the subject of the document.

```typescript
subject: string;
```

#### Property Value
- string

#### Remarks
[ [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

### template

Gets the template of the document.

```typescript
readonly template: string;
```

#### Property Value
- string

#### Remarks
[ [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

### title

Specifies the title of the document.

```typescript
title: string;
```

#### Property Value
- string

#### Remarks
[ [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

## Method Details

### load(options)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(options?: Word.Interfaces.DocumentPropertiesLoadOptions): Word.DocumentProperties;
```

#### Parameters
- options  
  [Word.Interfaces.DocumentPropertiesLoadOptions](/en-us/javascript/api/word/word.interfaces.documentpropertiesloadoptions)

Provides options for which properties of the object to load.

#### Returns
- [Word.DocumentProperties](/en-us/javascript/api/word/word.documentproperties)

### load(propertyNames)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.DocumentProperties;
```

#### Parameters
- propertyNames  
  string | string[]

A comma-delimited string or an array of strings that specify the properties to load.

#### Returns
- [Word.DocumentProperties](/en-us/javascript/api/word/word.documentproperties)

### load(propertyNamesAndPaths)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.DocumentProperties;
```

#### Parameters
- propertyNamesAndPaths  
```
{
select?: string;
expand?: string;
}
```

`propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

#### Returns
- [Word.DocumentProperties](/en-us/javascript/api/word/word.documentproperties)

### set(properties, options)

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.DocumentPropertiesUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

#### Parameters
- properties  
  [Word.Interfaces.DocumentPropertiesUpdateData](/en-us/javascript/api/word/word.interfaces.documentpropertiesupdatedata)

A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.

- options  
  [OfficeExtension.UpdateOptions](/en-us/javascript/api/office/officeextension.updateoptions)

Provides an option to suppress errors if the properties object tries to set any read-only properties.

#### Returns
- void

### set(properties)

Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.DocumentProperties): void;
```

#### Parameters
- properties  
  [Word.DocumentProperties](/en-us/javascript/api/word/word.documentproperties)

#### Returns
- void

### toJSON()

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.DocumentProperties` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.DocumentPropertiesData`) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.DocumentPropertiesData;
```

#### Returns
- [Word.Interfaces.DocumentPropertiesData](/en-us/javascript/api/word/word.interfaces.documentpropertiesdata)

### track()

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.DocumentProperties;
```

#### Returns
- [Word.DocumentProperties](/en-us/javascript/api/word/word.documentproperties)

### untrack()

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

```typescript
untrack(): Word.DocumentProperties;
```

#### Returns
- [Word.DocumentProperties](/en-us/javascript/api/word/word.documentproperties)