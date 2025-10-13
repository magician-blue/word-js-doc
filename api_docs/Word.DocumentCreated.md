# Word.DocumentCreated class

- Package: [word](/en-us/javascript/api/word)

The DocumentCreated object is the top level object created by Application.CreateDocument. A DocumentCreated object is a special Document object.

- Extends: [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks

[API set: WordApi 1.3]

#### Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/insert-external-document.yaml

// Updates the text of the current document with the text from another document passed in as a Base64-encoded string.
await Word.run(async (context) => {
  // Use the Base64-encoded string representation of the selected .docx file.
  const externalDoc: Word.DocumentCreated = context.application.createDocument(externalDocument);
  await context.sync();

  if (!Office.context.requirements.isSetSupported("WordApiHiddenDocument", "1.3")) {
    console.warn("The WordApiHiddenDocument 1.3 requirement set isn't supported on this client so can't proceed. Try this action on a platform that supports this requirement set.");
    return;
  }

  const externalDocBody: Word.Body = externalDoc.body;
  externalDocBody.load("text");
  await context.sync();

  // Insert the external document's text at the beginning of the current document's body.
  const externalDocBodyText = externalDocBody.text;
  const currentDocBody: Word.Body = context.document.body;
  currentDocBody.insertText(externalDocBodyText, Word.InsertLocation.start);
  await context.sync();
});
```

## Properties

- body — Gets the body object of the document. The body is the text that excludes headers, footers, footnotes, textboxes, etc.
- contentControls — Gets the collection of content control objects in the document. This includes content controls in the body of the document, headers, footers, textboxes, etc.
- context — The request context associated with the object. This connects the add-in's process to the Office host application's process.
- customXmlParts — Gets the custom XML parts in the document.
- properties — Gets the properties of the document.
- saved — Indicates whether the changes in the document have been saved. A value of true indicates that the document hasn't changed since it was saved.
- sections — Gets the collection of section objects in the document.
- settings — Gets the add-in's settings in the document.

## Methods

- addStyle(name, type) — Adds a style into the document by name and type.
- addStyle(name, type) — Adds a style into the document by name and type.
- deleteBookmark(name) — Deletes a bookmark, if it exists, from the document.
- getBookmarkRange(name) — Gets a bookmark's range. Throws an ItemNotFound error if the bookmark doesn't exist.
- getBookmarkRangeOrNullObject(name) — Gets a bookmark's range. If the bookmark doesn't exist, then this method will return an object with its isNullObject property set to true. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).
- getContentControls(options) — Gets the currently supported content controls in the document.
- getStyles() — Gets a StyleCollection object that represents the whole style set of the document.
- insertFileFromBase64(base64File, insertLocation, insertFileOptions) — Inserts a document into the target document at a specific location with additional properties. Headers, footers, watermarks, and other section properties are copied by default.
- load(options) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNames) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNamesAndPaths) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- open() — Opens the document.
- save(saveBehavior, fileName) — Saves the document.
- save(saveBehavior, fileName) — Saves the document.
- set(properties, options) — Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
- set(properties) — Sets multiple properties on the object at the same time, based on an existing loaded object.
- toJSON() — Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify().
- track() — Track the object for automatic adjustment based on surrounding changes in the document.
- untrack() — Release the memory associated with this object, if it has previously been tracked.

## Property Details

### body

Gets the body object of the document. The body is the text that excludes headers, footers, footnotes, textboxes, etc.

```typescript
readonly body: Word.Body;
```

- Property value: [Word.Body](/en-us/javascript/api/word/word.body)

Remarks

[API set: WordApiHiddenDocument 1.3]

---

### contentControls

Gets the collection of content control objects in the document. This includes content controls in the body of the document, headers, footers, textboxes, etc.

```typescript
readonly contentControls: Word.ContentControlCollection;
```

- Property value: [Word.ContentControlCollection](/en-us/javascript/api/word/word.contentcontrolcollection)

Remarks

[API set: WordApiHiddenDocument 1.3]

---

### context

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

- Property value: [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

---

### customXmlParts

Gets the custom XML parts in the document.

```typescript
readonly customXmlParts: Word.CustomXmlPartCollection;
```

- Property value: [Word.CustomXmlPartCollection](/en-us/javascript/api/word/word.customxmlpartcollection)

Remarks

[API set: WordApiHiddenDocument 1.4]

---

### properties

Gets the properties of the document.

```typescript
readonly properties: Word.DocumentProperties;
```

- Property value: [Word.DocumentProperties](/en-us/javascript/api/word/word.documentproperties)

Remarks

[API set: WordApiHiddenDocument 1.3]

---

### saved

Indicates whether the changes in the document have been saved. A value of true indicates that the document hasn't changed since it was saved.

```typescript
readonly saved: boolean;
```

- Property value: boolean

Remarks

[API set: WordApiHiddenDocument 1.3]

---

### sections

Gets the collection of section objects in the document.

```typescript
readonly sections: Word.SectionCollection;
```

- Property value: [Word.SectionCollection](/en-us/javascript/api/word/word.sectioncollection)

Remarks

[API set: WordApiHiddenDocument 1.3]

---

### settings

Gets the add-in's settings in the document.

```typescript
readonly settings: Word.SettingCollection;
```

- Property value: [Word.SettingCollection](/en-us/javascript/api/word/word.settingcollection)

Remarks

[API set: WordApiHiddenDocument 1.4]

## Method Details

### addStyle(name, type)

Adds a style into the document by name and type.

```typescript
addStyle(name: string, type: Word.StyleType): Word.Style;
```

Parameters

- name: string  
  Required. A string representing the style name.
- type: [Word.StyleType](/en-us/javascript/api/word/word.styletype)  
  Required. The style type, including character, list, paragraph, or table.

Returns

- [Word.Style](/en-us/javascript/api/word/word.style)

Remarks

[API set: WordApiHiddenDocument 1.5]

---

### addStyle(name, type)

Adds a style into the document by name and type.

```typescript
addStyle(name: string, type: "Character" | "List" | "Paragraph" | "Table"): Word.Style;
```

Parameters

- name: string  
  Required. A string representing the style name.
- type: "Character" | "List" | "Paragraph" | "Table"  
  Required. The style type, including character, list, paragraph, or table.

Returns

- [Word.Style](/en-us/javascript/api/word/word.style)

Remarks

[API set: WordApiHiddenDocument 1.5]

---

### deleteBookmark(name)

Deletes a bookmark, if it exists, from the document.

```typescript
deleteBookmark(name: string): void;
```

Parameters

- name: string  
  Required. The case-insensitive bookmark name.

Returns

- void

Remarks

[API set: WordApiHiddenDocument 1.4]

---

### getBookmarkRange(name)

Gets a bookmark's range. Throws an ItemNotFound error if the bookmark doesn't exist.

```typescript
getBookmarkRange(name: string): Word.Range;
```

Parameters

- name: string  
  Required. The case-insensitive bookmark name.

Returns

- [Word.Range](/en-us/javascript/api/word/word.range)

Remarks

[API set: WordApiHiddenDocument 1.4]

---

### getBookmarkRangeOrNullObject(name)

Gets a bookmark's range. If the bookmark doesn't exist, then this method will return an object with its isNullObject property set to true. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
getBookmarkRangeOrNullObject(name: string): Word.Range;
```

Parameters

- name: string  
  Required. The case-insensitive bookmark name. Only alphanumeric and underscore characters are supported. It must begin with a letter but if you want to tag the bookmark as hidden, then start the name with an underscore character. Names can't be longer than 40 characters.

Returns

- [Word.Range](/en-us/javascript/api/word/word.range)

Remarks

[API set: WordApiHiddenDocument 1.4]

---

### getContentControls(options)

Gets the currently supported content controls in the document.

```typescript
getContentControls(options?: Word.ContentControlOptions): Word.ContentControlCollection;
```

Parameters

- options: [Word.ContentControlOptions](/en-us/javascript/api/word/word.contentcontroloptions)  
  Optional. Options that define which content controls are returned.

Returns

- [Word.ContentControlCollection](/en-us/javascript/api/word/word.contentcontrolcollection)

Remarks

[API set: WordApiHiddenDocument 1.5]

Important: If specific types are provided in the options parameter, only content controls of supported types are returned. Be aware that an exception will be thrown on using methods of a generic [Word.ContentControl](/en-us/javascript/api/word/word.contentcontrol) that aren't relevant for the specific type. With time, additional types of content controls may be supported. Therefore, your add-in should request and handle specific types of content controls.

---

### getStyles()

Gets a StyleCollection object that represents the whole style set of the document.

```typescript
getStyles(): Word.StyleCollection;
```

Returns

- [Word.StyleCollection](/en-us/javascript/api/word/word.stylecollection)

Remarks

[API set: WordApiHiddenDocument 1.5]

---

### insertFileFromBase64(base64File, insertLocation, insertFileOptions)

Inserts a document into the target document at a specific location with additional properties. Headers, footers, watermarks, and other section properties are copied by default.

```typescript
insertFileFromBase64(
  base64File: string,
  insertLocation: Word.InsertLocation.replace | Word.InsertLocation.start | Word.InsertLocation.end | "Replace" | "Start" | "End",
  insertFileOptions?: Word.InsertFileOptions
): Word.SectionCollection;
```

Parameters

- base64File: string  
  Required. The Base64-encoded content of a .docx file.
- insertLocation: [replace](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-replace-member) | [start](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-start-member) | [end](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-end-member) | "Replace" | "Start" | "End"  
  Required. The value must be 'Replace', 'Start', or 'End'.
- insertFileOptions: [Word.InsertFileOptions](/en-us/javascript/api/word/word.insertfileoptions)  
  Optional. The additional properties that should be imported to the destination document.

Returns

- [Word.SectionCollection](/en-us/javascript/api/word/word.sectioncollection)

Remarks

[API set: WordApiHiddenDocument 1.5]

Insertion isn't supported if the document being inserted contains an ActiveX control (likely in a form field). Consider replacing such a form field with a content control or other option appropriate for your scenario.

---

### load(options)

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(options?: Word.Interfaces.DocumentCreatedLoadOptions): Word.DocumentCreated;
```

Parameters

- options: [Word.Interfaces.DocumentCreatedLoadOptions](/en-us/javascript/api/word/word.interfaces.documentcreatedloadoptions)  
  Provides options for which properties of the object to load.

Returns

- [Word.DocumentCreated](/en-us/javascript/api/word/word.documentcreated)

---

### load(propertyNames)

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.DocumentCreated;
```

Parameters

- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns

- [Word.DocumentCreated](/en-us/javascript/api/word/word.documentcreated)

---

### load(propertyNamesAndPaths)

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
  select?: string;
  expand?: string;
}): Word.DocumentCreated;
```

Parameters

- propertyNamesAndPaths: { select?: string; expand?: string; }  
  propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

Returns

- [Word.DocumentCreated](/en-us/javascript/api/word/word.documentcreated)

---

### open()

Opens the document.

```typescript
open(): void;
```

Returns

- void

Remarks

[API set: WordApi 1.3]

#### Examples

```typescript
// Create and open a new document in a new tab or window.
await Word.run(async (context) => {
  const externalDoc = context.application.createDocument();
  await context.sync();

  externalDoc.open();
  await context.sync();
});
```

---

### save(saveBehavior, fileName)

Saves the document.

```typescript
save(saveBehavior?: Word.SaveBehavior, fileName?: string): void;
```

Parameters

- saveBehavior: [Word.SaveBehavior](/en-us/javascript/api/word/word.savebehavior)  
  Optional. DocumentCreated only supports 'Save'.
- fileName: string  
  Optional. The file name (exclude file extension). Only takes effect for a new document.

Returns

- void

Remarks

[API set: WordApiHiddenDocument 1.3]

Note: The saveBehavior and fileName parameters were introduced in WordApiHiddenDocument 1.5.

---

### save(saveBehavior, fileName)

Saves the document.

```typescript
save(saveBehavior?: "Save" | "Prompt", fileName?: string): void;
```

Parameters

- saveBehavior: "Save" | "Prompt"  
  Optional. DocumentCreated only supports 'Save'.
- fileName: string  
  Optional. The file name (exclude file extension). Only takes effect for a new document.

Returns

- void

Remarks

[API set: WordApiHiddenDocument 1.3]

Note: The saveBehavior and fileName parameters were introduced in WordApiHiddenDocument 1.5.

---

### set(properties, options)

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.DocumentCreatedUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

Parameters

- properties: [Word.Interfaces.DocumentCreatedUpdateData](/en-us/javascript/api/word/word.interfaces.documentcreatedupdatedata)  
  A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
- options: [OfficeExtension.UpdateOptions](/en-us/javascript/api/office/officeextension.updateoptions)  
  Provides an option to suppress errors if the properties object tries to set any read-only properties.

Returns

- void

---

### set(properties)

Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.DocumentCreated): void;
```

Parameters

- properties: [Word.DocumentCreated](/en-us/javascript/api/word/word.documentcreated)

Returns

- void

---

### toJSON()

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.DocumentCreated object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.DocumentCreatedData) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.DocumentCreatedData;
```

Returns

- [Word.Interfaces.DocumentCreatedData](/en-us/javascript/api/word/word.interfaces.documentcreateddata)

---

### track()

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.DocumentCreated;
```

Returns

- [Word.DocumentCreated](/en-us/javascript/api/word/word.documentcreated)

---

### untrack()

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.DocumentCreated;
```

Returns

- [Word.DocumentCreated](/en-us/javascript/api/word/word.documentcreated)