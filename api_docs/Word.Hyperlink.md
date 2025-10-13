# Word.Hyperlink class

- Package: [word](https://learn.microsoft.com/en-us/javascript/api/word)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents a hyperlink in a Word document.

- Extends: [OfficeExtension.ClientObject](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientobject)

## Remarks

[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- address — Specifies the address (for example, a file name or URL) of the hyperlink.
- context — The request context associated with the object. This connects the add-in's process to the Office host application's process.
- emailSubject — Specifies the text string for the hyperlink's subject line.
- isExtraInfoRequired — Returns true if extra information is required to resolve the hyperlink.
- name — Returns the name of the Hyperlink object.
- range — Returns a Range object that represents the portion of the document that's contained within the hyperlink.
- screenTip — Specifies the text that appears as a ScreenTip when the mouse pointer is positioned over the hyperlink.
- subAddress — Specifies a named location in the destination of the hyperlink.
- target — Specifies the name of the frame or window in which to load the hyperlink.
- textToDisplay — Specifies the hyperlink's visible text in the document.
- type — Returns the hyperlink type.

## Methods

- addToFavorites() — Creates a shortcut to the document or hyperlink and adds it to the Favorites folder.
- createNewDocument(fileName, editNow, overwrite) — Creates a new document linked to the hyperlink.
- delete() — Deletes the hyperlink.
- load(options) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNames) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNamesAndPaths) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- set(properties, options) — Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
- set(properties) — Sets multiple properties on the object at the same time, based on an existing loaded object.
- toJSON() — Overrides the JavaScript toJSON() method to provide more useful output for JSON.stringify().
- track() — Track the object for automatic adjustment based on surrounding changes in the document.
- untrack() — Release the memory associated with this object, if it has previously been tracked.

## Property details

### address

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the address (for example, a file name or URL) of the hyperlink.

```typescript
address: string;
```

Property Value: string

Remarks

[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### context

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value: [Word.RequestContext](https://learn.microsoft.com/en-us/javascript/api/word/word.requestcontext)

---

### emailSubject

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the text string for the hyperlink's subject line.

```typescript
emailSubject: string;
```

Property Value: string

Remarks

[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### isExtraInfoRequired

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns true if extra information is required to resolve the hyperlink.

```typescript
readonly isExtraInfoRequired: boolean;
```

Property Value: boolean

Remarks

[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### name

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the name of the Hyperlink object.

```typescript
readonly name: string;
```

Property Value: string

Remarks

[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### range

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a Range object that represents the portion of the document that's contained within the hyperlink.

```typescript
readonly range: Word.Range;
```

Property Value: [Word.Range](https://learn.microsoft.com/en-us/javascript/api/word/word.range)

Remarks

[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### screenTip

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the text that appears as a ScreenTip when the mouse pointer is positioned over the hyperlink.

```typescript
screenTip: string;
```

Property Value: string

Remarks

[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### subAddress

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a named location in the destination of the hyperlink.

```typescript
subAddress: string;
```

Property Value: string

Remarks

[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### target

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the name of the frame or window in which to load the hyperlink.

```typescript
target: string;
```

Property Value: string

Remarks

[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### textToDisplay

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the hyperlink's visible text in the document.

```typescript
textToDisplay: string;
```

Property Value: string

Remarks

[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### type

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the hyperlink type.

```typescript
readonly type: Word.HyperlinkType | "Range" | "Shape" | "InlineShape";
```

Property Value: [Word.HyperlinkType](https://learn.microsoft.com/en-us/javascript/api/word/word.hyperlinktype) | "Range" | "Shape" | "InlineShape"

Remarks

[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Method details

### addToFavorites()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Creates a shortcut to the document or hyperlink and adds it to the Favorites folder.

```typescript
addToFavorites(): void;
```

Returns: void

Remarks

[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### createNewDocument(fileName, editNow, overwrite)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Creates a new document linked to the hyperlink.

```typescript
createNewDocument(fileName: string, editNow: boolean, overwrite: boolean): void;
```

Parameters

- fileName: string  
  Required. The name of the file.
- editNow: boolean  
  Required. true to start editing now.
- overwrite: boolean  
  Required. true to overwrite if there's another file with the same name.

Returns: void

Remarks

[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### delete()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Deletes the hyperlink.

```typescript
delete(): void;
```

Returns: void

Remarks

[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### load(options)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(options?: Word.Interfaces.HyperlinkLoadOptions): Word.Hyperlink;
```

Parameters

- options: [Word.Interfaces.HyperlinkLoadOptions](https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.hyperlinkloadoptions)  
  Provides options for which properties of the object to load.

Returns: [Word.Hyperlink](https://learn.microsoft.com/en-us/javascript/api/word/word.hyperlink)

---

### load(propertyNames)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.Hyperlink;
```

Parameters

- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns: [Word.Hyperlink](https://learn.microsoft.com/en-us/javascript/api/word/word.hyperlink)

---

### load(propertyNamesAndPaths)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
  select?: string;
  expand?: string;
}): Word.Hyperlink;
```

Parameters

- propertyNamesAndPaths:  
  select is a comma-delimited string that specifies the properties to load, and expand is a comma-delimited string that specifies the navigation properties to load.

Returns: [Word.Hyperlink](https://learn.microsoft.com/en-us/javascript/api/word/word.hyperlink)

---

### set(properties, options)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.HyperlinkUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

Parameters

- properties: [Word.Interfaces.HyperlinkUpdateData](https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.hyperlinkupdatedata)  
  A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
- options: [OfficeExtension.UpdateOptions](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.updateoptions)  
  Provides an option to suppress errors if the properties object tries to set any read-only properties.

Returns: void

---

### set(properties)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.Hyperlink): void;
```

Parameters

- properties: [Word.Hyperlink](https://learn.microsoft.com/en-us/javascript/api/word/word.hyperlink)

Returns: void

---

### toJSON()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.Hyperlink object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.HyperlinkData) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.HyperlinkData;
```

Returns: [Word.Interfaces.HyperlinkData](https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.hyperlinkdata)

---

### track()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.Hyperlink;
```

Returns: [Word.Hyperlink](https://learn.microsoft.com/en-us/javascript/api/word/word.hyperlink)

---

### untrack()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.Hyperlink;
```

Returns: [Word.Hyperlink](https://learn.microsoft.com/en-us/javascript/api/word/word.hyperlink)