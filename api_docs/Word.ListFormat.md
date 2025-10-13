# Word.ListFormat class

Package: [word](/en-us/javascript/api/word)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents the list formatting characteristics of a range.

Extends: [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks

[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- context  
  The request context associated with the object. This connects the add-in's process to the Office host application's process.

- isSingleList  
  Indicates whether the `ListFormat` object contains a single list.

- isSingleListTemplate  
  Indicates whether the `ListFormat` object contains a single list template.

- list  
  Returns a `List` object that represents the first formatted list contained in the `ListFormat` object.

- listLevelNumber  
  Specifies the list level number for the first paragraph for the `ListFormat` object.

- listString  
  Gets the string representation of the list value of the first paragraph in the range for the `ListFormat` object.

- listTemplate  
  Gets the list template associated with the `ListFormat` object.

- listType  
  Gets the type of the list for the `ListFormat` object.

- listValue  
  Gets the numeric value of the the first paragraph in the range for the `ListFormat` object.

## Methods

- applyBulletDefault(defaultListBehavior)  
  Adds bullets and formatting to the paragraphs in the range.

- applyBulletDefault(defaultListBehavior)  
  Adds bullets and formatting to the paragraphs in the range.

- applyListTemplateWithLevel(listTemplate, options)  
  Applies a list template with a specific level to the paragraphs in the range.

- applyNumberDefault(defaultListBehavior)  
  Adds numbering and formatting to the paragraphs in the range.

- applyNumberDefault(defaultListBehavior)  
  Adds numbering and formatting to the paragraphs in the range.

- applyOutlineNumberDefault(defaultListBehavior)  
  Adds outline numbering and formatting to the paragraphs in the range.

- applyOutlineNumberDefault(defaultListBehavior)  
  Adds outline numbering and formatting to the paragraphs in the range.

- canContinuePreviousList(listTemplate)  
  Determines whether the `ListFormat` object can continue a previous list.

- convertNumbersToText(numberType)  
  Converts numbers in the list to plain text.

- convertNumbersToText(numberType)  
  Converts numbers in the list to plain text.

- countNumberedItems(options)  
  Counts the numbered items in the list.

- listIndent()  
  Indents the list by one level.

- listOutdent()  
  Outdents the list by one level.

- load(options)  
  Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

- load(propertyNames)  
  Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

- load(propertyNamesAndPaths)  
  Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

- removeNumbers(numberType)  
  Removes numbering from the list.

- removeNumbers(numberType)  
  Removes numbering from the list.

- set(properties, options)  
  Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

- set(properties)  
  Sets multiple properties on the object at the same time, based on an existing loaded object.

- toJSON()  
  Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`.

- track()  
  Track the object for automatic adjustment based on surrounding changes in the document.

- untrack()  
  Release the memory associated with this object, if it has previously been tracked.

## Property Details

### context

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

#### Property value
[Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

---

### isSingleList

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Indicates whether the `ListFormat` object contains a single list.

```typescript
readonly isSingleList: boolean;
```

#### Property value
boolean

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### isSingleListTemplate

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Indicates whether the `ListFormat` object contains a single list template.

```typescript
readonly isSingleListTemplate: boolean;
```

#### Property value
boolean

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### list

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `List` object that represents the first formatted list contained in the `ListFormat` object.

```typescript
readonly list: Word.List;
```

#### Property value
[Word.List](/en-us/javascript/api/word/word.list)

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### listLevelNumber

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the list level number for the first paragraph for the `ListFormat` object.

```typescript
listLevelNumber: number;
```

#### Property value
number

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### listString

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the string representation of the list value of the first paragraph in the range for the `ListFormat` object.

```typescript
readonly listString: string;
```

#### Property value
string

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### listTemplate

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the list template associated with the `ListFormat` object.

```typescript
readonly listTemplate: Word.ListTemplate;
```

#### Property value
[Word.ListTemplate](/en-us/javascript/api/word/word.listtemplate)

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### listType

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the type of the list for the `ListFormat` object.

```typescript
readonly listType: Word.ListType | "ListNoNumbering" | "ListListNumOnly" | "ListBullet" | "ListSimpleNumbering" | "ListOutlineNumbering" | "ListMixedNumbering" | "ListPictureBullet";
```

#### Property value
[Word.ListType](/en-us/javascript/api/word/word.listtype) | "ListNoNumbering" | "ListListNumOnly" | "ListBullet" | "ListSimpleNumbering" | "ListOutlineNumbering" | "ListMixedNumbering" | "ListPictureBullet"

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### listValue

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the numeric value of the the first paragraph in the range for the `ListFormat` object.

```typescript
readonly listValue: number;
```

#### Property value
number

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Method Details

### applyBulletDefault(defaultListBehavior)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Adds bullets and formatting to the paragraphs in the range.

```typescript
applyBulletDefault(defaultListBehavior: Word.DefaultListBehavior): void;
```

#### Parameters
- defaultListBehavior: [Word.DefaultListBehavior](/en-us/javascript/api/word/word.defaultlistbehavior)  
  Optional. Specifies the default list behavior. Default is `DefaultListBehavior.word97`.

#### Returns
void

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### applyBulletDefault(defaultListBehavior)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Adds bullets and formatting to the paragraphs in the range.

```typescript
applyBulletDefault(defaultListBehavior: "Word97" | "Word2000" | "Word2002"): void;
```

#### Parameters
- defaultListBehavior: "Word97" | "Word2000" | "Word2002"  
  Optional. Specifies the default list behavior. Default is `DefaultListBehavior.word97`.

#### Returns
void

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### applyListTemplateWithLevel(listTemplate, options)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Applies a list template with a specific level to the paragraphs in the range.

```typescript
applyListTemplateWithLevel(listTemplate: Word.ListTemplate, options?: Word.ListTemplateApplyOptions): void;
```

#### Parameters
- listTemplate: [Word.ListTemplate](/en-us/javascript/api/word/word.listtemplate)  
  The list template to apply.

- options: [Word.ListTemplateApplyOptions](/en-us/javascript/api/word/word.listtemplateapplyoptions)  
  Optional. Options for applying the list template, such as whether to continue the previous list or which part of the list to apply the template to.

#### Returns
void

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### applyNumberDefault(defaultListBehavior)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Adds numbering and formatting to the paragraphs in the range.

```typescript
applyNumberDefault(defaultListBehavior: Word.DefaultListBehavior): void;
```

#### Parameters
- defaultListBehavior: [Word.DefaultListBehavior](/en-us/javascript/api/word/word.defaultlistbehavior)  
  Optional. Specifies the default list behavior.

#### Returns
void

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### applyNumberDefault(defaultListBehavior)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Adds numbering and formatting to the paragraphs in the range.

```typescript
applyNumberDefault(defaultListBehavior: "Word97" | "Word2000" | "Word2002"): void;
```

#### Parameters
- defaultListBehavior: "Word97" | "Word2000" | "Word2002"  
  Optional. Specifies the default list behavior.

#### Returns
void

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### applyOutlineNumberDefault(defaultListBehavior)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Adds outline numbering and formatting to the paragraphs in the range.

```typescript
applyOutlineNumberDefault(defaultListBehavior: Word.DefaultListBehavior): void;
```

#### Parameters
- defaultListBehavior: [Word.DefaultListBehavior](/en-us/javascript/api/word/word.defaultlistbehavior)  
  Optional. Specifies the default list behavior.

#### Returns
void

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### applyOutlineNumberDefault(defaultListBehavior)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Adds outline numbering and formatting to the paragraphs in the range.

```typescript
applyOutlineNumberDefault(defaultListBehavior: "Word97" | "Word2000" | "Word2002"): void;
```

#### Parameters
- defaultListBehavior: "Word97" | "Word2000" | "Word2002"  
  Optional. Specifies the default list behavior.

#### Returns
void

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### canContinuePreviousList(listTemplate)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Determines whether the `ListFormat` object can continue a previous list.

```typescript
canContinuePreviousList(listTemplate: Word.ListTemplate): OfficeExtension.ClientResult<Word.Continue>;
```

#### Parameters
- listTemplate: [Word.ListTemplate](/en-us/javascript/api/word/word.listtemplate)  
  The list template to check.

#### Returns
[OfficeExtension.ClientResult](/en-us/javascript/api/office/officeextension.clientresult)<[Word.Continue](/en-us/javascript/api/word/word.continue)>

A `Continue` value indicating whether continuation is possible.

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### convertNumbersToText(numberType)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Converts numbers in the list to plain text.

```typescript
convertNumbersToText(numberType: Word.NumberType): void;
```

#### Parameters
- numberType: [Word.NumberType](/en-us/javascript/api/word/word.numbertype)  
  Optional. The type of number to convert.

#### Returns
void

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### convertNumbersToText(numberType)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Converts numbers in the list to plain text.

```typescript
convertNumbersToText(numberType: "Paragraph" | "ListNum" | "AllNumbers"): void;
```

#### Parameters
- numberType: "Paragraph" | "ListNum" | "AllNumbers"  
  Optional. The type of number to convert.

#### Returns
void

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### countNumberedItems(options)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Counts the numbered items in the list.

```typescript
countNumberedItems(options?: Word.ListFormatCountNumberedItemsOptions): OfficeExtension.ClientResult<number>;
```

#### Parameters
- options: [Word.ListFormatCountNumberedItemsOptions](/en-us/javascript/api/word/word.listformatcountnumbereditemsoptions)  
  Optional. Options for counting numbered items, such as the type of number and the level to count.

#### Returns
[OfficeExtension.ClientResult](/en-us/javascript/api/office/officeextension.clientresult)<number>

The number of items.

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### listIndent()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Indents the list by one level.

```typescript
listIndent(): void;
```

#### Returns
void

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### listOutdent()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Outdents the list by one level.

```typescript
listOutdent(): void;
```

#### Returns
void

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### load(options)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(options?: Word.Interfaces.ListFormatLoadOptions): Word.ListFormat;
```

#### Parameters
- options: [Word.Interfaces.ListFormatLoadOptions](/en-us/javascript/api/word/word.interfaces.listformatloadoptions)  
  Provides options for which properties of the object to load.

#### Returns
[Word.ListFormat](/en-us/javascript/api/word/word.listformat)

---

### load(propertyNames)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.ListFormat;
```

#### Parameters
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

#### Returns
[Word.ListFormat](/en-us/javascript/api/word/word.listformat)

---

### load(propertyNamesAndPaths)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
    select?: string;
    expand?: string;
}): Word.ListFormat;
```

#### Parameters
- propertyNamesAndPaths:  
  `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

#### Returns
[Word.ListFormat](/en-us/javascript/api/word/word.listformat)

---

### removeNumbers(numberType)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Removes numbering from the list.

```typescript
removeNumbers(numberType: Word.NumberType): void;
```

#### Parameters
- numberType: [Word.NumberType](/en-us/javascript/api/word/word.numbertype)  
  Optional. The type of number to remove.

#### Returns
void

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### removeNumbers(numberType)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Removes numbering from the list.

```typescript
removeNumbers(numberType: "Paragraph" | "ListNum" | "AllNumbers"): void;
```

#### Parameters
- numberType: "Paragraph" | "ListNum" | "AllNumbers"  
  Optional. The type of number to remove.

#### Returns
void

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### set(properties, options)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.ListFormatUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

#### Parameters
- properties: [Word.Interfaces.ListFormatUpdateData](/en-us/javascript/api/word/word.interfaces.listformatupdatedata)  
  A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.

- options: [OfficeExtension.UpdateOptions](/en-us/javascript/api/office/officeextension.updateoptions)  
  Provides an option to suppress errors if the properties object tries to set any read-only properties.

#### Returns
void

---

### set(properties)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.ListFormat): void;
```

#### Parameters
- properties: [Word.ListFormat](/en-us/javascript/api/word/word.listformat)

#### Returns
void

---

### toJSON()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.ListFormat` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.ListFormatData`) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.ListFormatData;
```

#### Returns
[Word.Interfaces.ListFormatData](/en-us/javascript/api/word/word.interfaces.listformatdata)

---

### track()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.ListFormat;
```

#### Returns
[Word.ListFormat](/en-us/javascript/api/word/word.listformat)

---

### untrack()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

```typescript
untrack(): Word.ListFormat;
```

#### Returns
[Word.ListFormat](/en-us/javascript/api/word/word.listformat)