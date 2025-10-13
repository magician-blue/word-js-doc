# Word.PictureContentControl class

Package: [word](/en-us/javascript/api/word)

> Note
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents the PictureContentControl object.

Extends: [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

## Properties

| Property | Description |
| --- | --- |
| [appearance](#appearance) | Specifies the appearance of the content control. |
| [color](#color) | Specifies the red-green-blue (RGB) value of the color of the content control. You can provide the value in the '#RRGGBB' format. |
| [context](#context) | The request context associated with the object. This connects the add-in's process to the Office host application's process. |
| [id](#id) | Returns the identification for the content control. |
| [isTemporary](#istemporary) | Specifies whether to remove the content control from the active document when the user edits the contents of the control. |
| [level](#level) | Returns the level of the content control—whether the content control surrounds text, paragraphs, table cells, or table rows; or if it is inline. |
| [lockContentControl](#lockcontentcontrol) | Specifies if the content control is locked (can't be deleted). true means that the user can't delete it from the active document, false means it can be deleted. |
| [lockContents](#lockcontents) | Specifies if the contents of the content control are locked (not editable). true means the user can't edit the contents, false means the contents are editable. |
| [placeholderText](#placeholdertext) | Returns a BuildingBlock object that represents the placeholder text for the content control. |
| [range](#range) | Returns a Range object that represents the contents of the content control in the active document. |
| [showingPlaceholderText](#showingplaceholdertext) | Returns whether the placeholder text for the content control is being displayed. |
| [tag](#tag) | Specifies a tag to identify the content control. |
| [title](#title) | Specifies the title for the content control. |
| [xmlMapping](#xmlmapping) | Returns an XmlMapping object that represents the mapping of the content control to XML data in the data store of the document. |

## Methods

| Method | Description |
| --- | --- |
| [copy()](#copy) | Copies the content control from the active document to the Clipboard. |
| [cut()](#cut) | Removes the content control from the active document and moves the content control to the Clipboard. |
| [delete(deleteContents)](#delete) | Deletes the content control and optionally its contents. |
| [load(options)](#loadoptions) | Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties. |
| [load(propertyNames)](#loadpropertynames) | Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties. |
| [load(propertyNamesAndPaths)](#loadpropertynamesandpaths) | Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties. |
| [set(properties, options)](#setproperties-options) | Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type. |
| [set(properties)](#setproperties) | Sets multiple properties on the object at the same time, based on an existing loaded object. |
| [setPlaceholderText(options)](#setplaceholdertext) | Sets the placeholder text that displays in the content control until a user enters their own text. |
| [toJSON()](#tojson) | Overrides the JavaScript toJSON() method to provide more useful output when an API object is passed to JSON.stringify(). |
| [track()](#track) | Track the object for automatic adjustment based on surrounding changes in the document. |
| [untrack()](#untrack) | Release the memory associated with this object, if it has previously been tracked. |

---

## Property Details

### appearance

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the appearance of the content control.

```typescript
appearance: Word.ContentControlAppearance | "BoundingBox" | "Tags" | "Hidden";
```

Property Value: [Word.ContentControlAppearance](/en-us/javascript/api/word/word.contentcontrolappearance) | "BoundingBox" | "Tags" | "Hidden"

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

### color

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the red-green-blue (RGB) value of the color of the content control. You can provide the value in the '#RRGGBB' format.

```typescript
color: string;
```

Property Value: string

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

### context

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value: [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### id

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the identification for the content control.

```typescript
readonly id: string;
```

Property Value: string

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

### isTemporary

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether to remove the content control from the active document when the user edits the contents of the control.

```typescript
isTemporary: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

### level

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the level of the content control—whether the content control surrounds text, paragraphs, table cells, or table rows; or if it is inline.

```typescript
readonly level: Word.ContentControlLevel | "Inline" | "Paragraph" | "Row" | "Cell";
```

Property Value: [Word.ContentControlLevel](/en-us/javascript/api/word/word.contentcontrollevel) | "Inline" | "Paragraph" | "Row" | "Cell"

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

### lockContentControl

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the content control is locked (can't be deleted). true means that the user can't delete it from the active document, false means it can be deleted.

```typescript
lockContentControl: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

### lockContents

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the contents of the content control are locked (not editable). true means the user can't edit the contents, false means the contents are editable.

```typescript
lockContents: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

### placeholderText

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a BuildingBlock object that represents the placeholder text for the content control.

```typescript
readonly placeholderText: Word.BuildingBlock;
```

Property Value: [Word.BuildingBlock](/en-us/javascript/api/word/word.buildingblock)

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

### range

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a Range object that represents the contents of the content control in the active document.

```typescript
readonly range: Word.Range;
```

Property Value: [Word.Range](/en-us/javascript/api/word/word.range)

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

### showingPlaceholderText

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns whether the placeholder text for the content control is being displayed.

```typescript
readonly showingPlaceholderText: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

### tag

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a tag to identify the content control.

```typescript
tag: string;
```

Property Value: string

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

### title

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the title for the content control.

```typescript
title: string;
```

Property Value: string

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

### xmlMapping

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns an XmlMapping object that represents the mapping of the content control to XML data in the data store of the document.

```typescript
readonly xmlMapping: Word.XmlMapping;
```

Property Value: [Word.XmlMapping](/en-us/javascript/api/word/word.xmlmapping)

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

---

## Method Details

### copy()

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Copies the content control from the active document to the Clipboard.

```typescript
copy(): void;
```

Returns: void

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

### cut()

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Removes the content control from the active document and moves the content control to the Clipboard.

```typescript
cut(): void;
```

Returns: void

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

### delete(deleteContents)

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Deletes the content control and optionally its contents.

```typescript
delete(deleteContents?: boolean): void;
```

Parameters:
- deleteContents: boolean  
  Optional. Decides whether to delete the contents of the content control.

Returns: void

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

### load(options)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(options?: Word.Interfaces.PictureContentControlLoadOptions): Word.PictureContentControl;
```

Parameters:
- options: [Word.Interfaces.PictureContentControlLoadOptions](/en-us/javascript/api/word/word.interfaces.picturecontentcontrolloadoptions)  
  Provides options for which properties of the object to load.

Returns: [Word.PictureContentControl](/en-us/javascript/api/word/word.picturecontentcontrol)

### load(propertyNames)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.PictureContentControl;
```

Parameters:
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns: [Word.PictureContentControl](/en-us/javascript/api/word/word.picturecontentcontrol)

### load(propertyNamesAndPaths)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.PictureContentControl;
```

Parameters:
- propertyNamesAndPaths: `{ select?: string; expand?: string; }`  
  `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

Returns: [Word.PictureContentControl](/en-us/javascript/api/word/word.picturecontentcontrol)

### set(properties, options)

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.PictureContentControlUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

Parameters:
- properties: [Word.Interfaces.PictureContentControlUpdateData](/en-us/javascript/api/word/word.interfaces.picturecontentcontrolupdatedata)  
  A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
- options: [OfficeExtension.UpdateOptions](/en-us/javascript/api/office/officeextension.updateoptions)  
  Provides an option to suppress errors if the properties object tries to set any read-only properties.

Returns: void

### set(properties)

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.PictureContentControl): void;
```

Parameters:
- properties: [Word.PictureContentControl](/en-us/javascript/api/word/word.picturecontentcontrol)

Returns: void

### setPlaceholderText(options)

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets the placeholder text that displays in the content control until a user enters their own text.

```typescript
setPlaceholderText(options?: Word.ContentControlPlaceholderOptions): void;
```

Parameters:
- options: [Word.ContentControlPlaceholderOptions](/en-us/javascript/api/word/word.contentcontrolplaceholderoptions)  
  Optional. The options for configuring the content control's placeholder text.

Returns: void

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

### toJSON()

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.PictureContentControl` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.PictureContentControlData`) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.PictureContentControlData;
```

Returns: [Word.Interfaces.PictureContentControlData](/en-us/javascript/api/word/word.interfaces.picturecontentcontroldata)

### track()

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.PictureContentControl;
```

Returns: [Word.PictureContentControl](/en-us/javascript/api/word/word.picturecontentcontrol)

### untrack()

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

```typescript
untrack(): Word.PictureContentControl;
```

Returns: [Word.PictureContentControl](/en-us/javascript/api/word/word.picturecontentcontrol)