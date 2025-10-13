# Word.GroupContentControl class

Package: [word](/en-us/javascript/api/word)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents the GroupContentControl object.

Extends: [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks

[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- appearance — Specifies the appearance of the content control.
- color — Specifies the red-green-blue (RGB) value of the color of the content control. You can provide the value in the '#RRGGBB' format.
- context — The request context associated with the object. This connects the add-in's process to the Office host application's process.
- id — Returns the identification for the content control.
- isTemporary — Specifies whether to remove the content control from the active document when the user edits the contents of the control.
- level — Gets the level of the content control—whether the content control surrounds text, paragraphs, table cells, or table rows; or if it is inline.
- lockContentControl — Specifies if the content control is locked (can't be deleted). true means that the user can't delete it from the active document, false means it can be deleted.
- lockContents — Specifies if the contents of the content control are locked (not editable). true means the user can't edit the contents, false means the contents are editable.
- placeholderText — Returns a BuildingBlock object that represents the placeholder text for the content control.
- range — Gets a Range object that represents the contents of the content control in the active document.
- showingPlaceholderText — Returns whether the placeholder text for the content control is being displayed.
- tag — Specifies a tag to identify the content control.
- title — Specifies the title for the content control.
- xmlMapping — Gets an XmlMapping object that represents the mapping of the content control to XML data in the data store of the document.

## Methods

- copy() — Copies the content control from the active document to the Clipboard.
- cut() — Removes the content control from the active document and moves the content control to the Clipboard.
- delete(deleteContents) — Deletes the content control and optionally its contents.
- load(options) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNames) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNamesAndPaths) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- set(properties, options) — Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
- set(properties) — Sets multiple properties on the object at the same time, based on an existing loaded object.
- setPlaceholderText(options) — Sets the placeholder text that displays in the content control until a user enters their own text.
- toJSON() — Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). Whereas the original Word.GroupContentControl object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.GroupContentControlData) that contains shallow copies of any loaded child properties from the original object.
- track() — Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
- ungroup() — Removes the group content control from the document so that its child content controls are no longer nested and can be freely edited.
- untrack() — Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

## Property Details

### appearance

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the appearance of the content control.

```typescript
appearance: Word.ContentControlAppearance | "BoundingBox" | "Tags" | "Hidden";
```

#### Property Value

[Word.ContentControlAppearance](/en-us/javascript/api/word/word.contentcontrolappearance) | "BoundingBox" | "Tags" | "Hidden"

#### Remarks

[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### color

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the red-green-blue (RGB) value of the color of the content control. You can provide the value in the '#RRGGBB' format.

```typescript
color: string;
```

#### Property Value

string

#### Remarks

[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### context

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

#### Property Value

[Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

---

### id

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the identification for the content control.

```typescript
readonly id: string;
```

#### Property Value

string

#### Remarks

[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### isTemporary

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether to remove the content control from the active document when the user edits the contents of the control.

```typescript
isTemporary: boolean;
```

#### Property Value

boolean

#### Remarks

[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### level

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the level of the content control—whether the content control surrounds text, paragraphs, table cells, or table rows; or if it is inline.

```typescript
readonly level: Word.ContentControlLevel | "Inline" | "Paragraph" | "Row" | "Cell";
```

#### Property Value

[Word.ContentControlLevel](/en-us/javascript/api/word/word.contentcontrollevel) | "Inline" | "Paragraph" | "Row" | "Cell"

#### Remarks

[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### lockContentControl

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the content control is locked (can't be deleted). true means that the user can't delete it from the active document, false means it can be deleted.

```typescript
lockContentControl: boolean;
```

#### Property Value

boolean

#### Remarks

[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### lockContents

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the contents of the content control are locked (not editable). true means the user can't edit the contents, false means the contents are editable.

```typescript
lockContents: boolean;
```

#### Property Value

boolean

#### Remarks

[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### placeholderText

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a BuildingBlock object that represents the placeholder text for the content control.

```typescript
readonly placeholderText: Word.BuildingBlock;
```

#### Property Value

[Word.BuildingBlock](/en-us/javascript/api/word/word.buildingblock)

#### Remarks

[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### range

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets a Range object that represents the contents of the content control in the active document.

```typescript
readonly range: Word.Range;
```

#### Property Value

[Word.Range](/en-us/javascript/api/word/word.range)

#### Remarks

[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### showingPlaceholderText

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns whether the placeholder text for the content control is being displayed.

```typescript
readonly showingPlaceholderText: boolean;
```

#### Property Value

boolean

#### Remarks

[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### tag

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a tag to identify the content control.

```typescript
tag: string;
```

#### Property Value

string

#### Remarks

[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### title

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the title for the content control.

```typescript
title: string;
```

#### Property Value

string

#### Remarks

[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### xmlMapping

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets an XmlMapping object that represents the mapping of the content control to XML data in the data store of the document.

```typescript
readonly xmlMapping: Word.XmlMapping;
```

#### Property Value

[Word.XmlMapping](/en-us/javascript/api/word/word.xmlmapping)

#### Remarks

[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Method Details

### copy()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Copies the content control from the active document to the Clipboard.

```typescript
copy(): void;
```

#### Returns

void

#### Remarks

[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### cut()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Removes the content control from the active document and moves the content control to the Clipboard.

```typescript
cut(): void;
```

#### Returns

void

#### Remarks

[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### delete(deleteContents)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Deletes the content control and optionally its contents.

```typescript
delete(deleteContents: boolean): void;
```

#### Parameters

- deleteContents: boolean

Optional. Whether to delete the contents inside the control.

#### Returns

void

#### Remarks

[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### load(options)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(options?: Word.Interfaces.GroupContentControlLoadOptions): Word.GroupContentControl;
```

#### Parameters

- options: [Word.Interfaces.GroupContentControlLoadOptions](/en-us/javascript/api/word/word.interfaces.groupcontentcontrolloadoptions)

Provides options for which properties of the object to load.

#### Returns

[Word.GroupContentControl](/en-us/javascript/api/word/word.groupcontentcontrol)

---

### load(propertyNames)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.GroupContentControl;
```

#### Parameters

- propertyNames: string | string[]

A comma-delimited string or an array of strings that specify the properties to load.

#### Returns

[Word.GroupContentControl](/en-us/javascript/api/word/word.groupcontentcontrol)

---

### load(propertyNamesAndPaths)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.GroupContentControl;
```

#### Parameters

- propertyNamesAndPaths: { select?: string; expand?: string; }

propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

#### Returns

[Word.GroupContentControl](/en-us/javascript/api/word/word.groupcontentcontrol)

---

### set(properties, options)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.GroupContentControlUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

#### Parameters

- properties: [Word.Interfaces.GroupContentControlUpdateData](/en-us/javascript/api/word/word.interfaces.groupcontentcontrolupdatedata)

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
set(properties: Word.GroupContentControl): void;
```

#### Parameters

- properties: [Word.GroupContentControl](/en-us/javascript/api/word/word.groupcontentcontrol)

#### Returns

void

---

### setPlaceholderText(options)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets the placeholder text that displays in the content control until a user enters their own text.

```typescript
setPlaceholderText(options?: Word.ContentControlPlaceholderOptions): void;
```

#### Parameters

- options: [Word.ContentControlPlaceholderOptions](/en-us/javascript/api/word/word.contentcontrolplaceholderoptions)

Optional. The options for configuring the content control's placeholder text.

#### Returns

void

#### Remarks

[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### toJSON()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.GroupContentControl object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.GroupContentControlData) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.GroupContentControlData;
```

#### Returns

[Word.Interfaces.GroupContentControlData](/en-us/javascript/api/word/word.interfaces.groupcontentcontroldata)

---

### track()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.GroupContentControl;
```

#### Returns

[Word.GroupContentControl](/en-us/javascript/api/word/word.groupcontentcontrol)

---

### ungroup()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Removes the group content control from the document so that its child content controls are no longer nested and can be freely edited.

```typescript
ungroup(): void;
```

#### Returns

void

#### Remarks

[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### untrack()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.GroupContentControl;
```

#### Returns

[Word.GroupContentControl](/en-us/javascript/api/word/word.groupcontentcontrol)