# Word.OleFormat class

Package: [word](/en-us/javascript/api/word)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents the OLE characteristics (other than linking) for an OLE object, ActiveX control, or field.

Extends
- [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks

[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- [classType](#word-word-oleformat-classtype-member) — Specifies the class type for the specified OLE object, picture, or field.
- [context](#word-word-oleformat-context-member) — The request context associated with the object. This connects the add-in's process to the Office host application's process.
- [iconIndex](#word-word-oleformat-iconindex-member) — Specifies the icon that is used when the `displayAsIcon` property is `true`.
- [iconLabel](#word-word-oleformat-iconlabel-member) — Specifies the text displayed below the icon for the OLE object.
- [iconName](#word-word-oleformat-iconname-member) — Specifies the program file in which the icon for the OLE object is stored.
- [iconPath](#word-word-oleformat-iconpath-member) — Gets the path of the file in which the icon for the OLE object is stored.
- [isDisplayedAsIcon](#word-word-oleformat-isdisplayedasicon-member) — Gets whether the specified object is displayed as an icon.
- [isFormattingPreservedOnUpdate](#word-word-oleformat-isformattingpreservedonupdate-member) — Specifies whether formatting done in Microsoft Word to the linked OLE object is preserved.
- [label](#word-word-oleformat-label-member) — Gets a string that's used to identify the portion of the source file that's being linked.
- [progID](#word-word-oleformat-progid-member) — Gets the programmatic identifier (`ProgId`) for the specified OLE object.

## Methods

- [activate()](#word-word-oleformat-activate-member1) — Activates the `OleFormat` object.
- [activateAs(classType)](#word-word-oleformat-activateas-member1) — Sets the Windows registry value that determines the default application used to activate the specified OLE object.
- [doVerb(verbIndex)](#word-word-oleformat-doverb-member1) — Requests that the OLE object perform one of its available verbs.
- [doVerb(verbIndex)](#word-word-oleformat-doverb-member2) — Requests that the OLE object perform one of its available verbs.
- [edit()](#word-word-oleformat-edit-member1) — Opens the OLE object for editing in the application it was created in.
- [load(options)](#word-word-oleformat-load-member1) — Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- [load(propertyNames)](#word-word-oleformat-load-member2) — Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- [load(propertyNamesAndPaths)](#word-word-oleformat-load-member3) — Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- [open()](#word-word-oleformat-open-member1) — Opens the `OleFormat` object.
- [set(properties, options)](#word-word-oleformat-set-member1) — Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
- [set(properties)](#word-word-oleformat-set-member2) — Sets multiple properties on the object at the same time, based on an existing loaded object.
- [toJSON()](#word-word-oleformat-tojson-member1) — Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`.
- [track()](#word-word-oleformat-track-member1) — Track the object for automatic adjustment based on surrounding changes in the document.
- [untrack()](#word-word-oleformat-untrack-member1) — Release the memory associated with this object, if it has previously been tracked.

## Property Details

### classType
Id: word-word-oleformat-classtype-member

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the class type for the specified OLE object, picture, or field.

```typescript
classType: string;
```

Property Value
- string

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### context
Id: word-word-oleformat-context-member

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value
- [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### iconIndex
Id: word-word-oleformat-iconindex-member

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the icon that is used when the `displayAsIcon` property is `true`.

```typescript
iconIndex: number;
```

Property Value
- number

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### iconLabel
Id: word-word-oleformat-iconlabel-member

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the text displayed below the icon for the OLE object.

```typescript
iconLabel: string;
```

Property Value
- string

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### iconName
Id: word-word-oleformat-iconname-member

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the program file in which the icon for the OLE object is stored.

```typescript
iconName: string;
```

Property Value
- string

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### iconPath
Id: word-word-oleformat-iconpath-member

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the path of the file in which the icon for the OLE object is stored.

```typescript
readonly iconPath: string;
```

Property Value
- string

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### isDisplayedAsIcon
Id: word-word-oleformat-isdisplayedasicon-member

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether the specified object is displayed as an icon.

```typescript
readonly isDisplayedAsIcon: boolean;
```

Property Value
- boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### isFormattingPreservedOnUpdate
Id: word-word-oleformat-isformattingpreservedonupdate-member

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether formatting done in Microsoft Word to the linked OLE object is preserved.

```typescript
isFormattingPreservedOnUpdate: boolean;
```

Property Value
- boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### label
Id: word-word-oleformat-label-member

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets a string that's used to identify the portion of the source file that's being linked.

```typescript
readonly label: string;
```

Property Value
- string

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### progID
Id: word-word-oleformat-progid-member

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the programmatic identifier (`ProgId`) for the specified OLE object.

```typescript
readonly progID: string;
```

Property Value
- string

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Method Details

### activate()
Id: word-word-oleformat-activate-member(1)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Activates the `OleFormat` object.

```typescript
activate(): void;
```

Returns
- void

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### activateAs(classType)
Id: word-word-oleformat-activateas-member(1)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets the Windows registry value that determines the default application used to activate the specified OLE object.

```typescript
activateAs(classType: string): void;
```

Parameters
- classType: string  
  The class type to activate as.

Returns
- void

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### doVerb(verbIndex)
Id: word-word-oleformat-doverb-member(1)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Requests that the OLE object perform one of its available verbs.

```typescript
doVerb(verbIndex: Word.OleVerb): void;
```

Parameters
- verbIndex: [Word.OleVerb](/en-us/javascript/api/word/word.oleverb)  
  Optional. The index of the verb to perform.

Returns
- void

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### doVerb(verbIndex)
Id: word-word-oleformat-doverb-member(2)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Requests that the OLE object perform one of its available verbs.

```typescript
doVerb(verbIndex: "Primary" | "Show" | "Open" | "Hide" | "UiActivate" | "InPlaceActivate" | "DiscardUndoState"): void;
```

Parameters
- verbIndex: "Primary" | "Show" | "Open" | "Hide" | "UiActivate" | "InPlaceActivate" | "DiscardUndoState"  
  Optional. The index of the verb to perform.

Returns
- void

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### edit()
Id: word-word-oleformat-edit-member(1)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Opens the OLE object for editing in the application it was created in.

```typescript
edit(): void;
```

Returns
- void

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### load(options)
Id: word-word-oleformat-load-member(1)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(options?: Word.Interfaces.OleFormatLoadOptions): Word.OleFormat;
```

Parameters
- options: [Word.Interfaces.OleFormatLoadOptions](/en-us/javascript/api/word/word.interfaces.oleformatloadoptions)  
  Provides options for which properties of the object to load.

Returns
- [Word.OleFormat](/en-us/javascript/api/word/word.oleformat)

### load(propertyNames)
Id: word-word-oleformat-load-member(2)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.OleFormat;
```

Parameters
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns
- [Word.OleFormat](/en-us/javascript/api/word/word.oleformat)

### load(propertyNamesAndPaths)
Id: word-word-oleformat-load-member(3)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.OleFormat;
```

Parameters
- propertyNamesAndPaths: { select?: string; expand?: string; }  
  `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

Returns
- [Word.OleFormat](/en-us/javascript/api/word/word.oleformat)

### open()
Id: word-word-oleformat-open-member(1)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Opens the `OleFormat` object.

```typescript
open(): void;
```

Returns
- void

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### set(properties, options)
Id: word-word-oleformat-set-member(1)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.OleFormatUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

Parameters
- properties: [Word.Interfaces.OleFormatUpdateData](/en-us/javascript/api/word/word.interfaces.oleformatupdatedata)  
  A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
- options: [OfficeExtension.UpdateOptions](/en-us/javascript/api/office/officeextension.updateoptions)  
  Provides an option to suppress errors if the properties object tries to set any read-only properties.

Returns
- void

### set(properties)
Id: word-word-oleformat-set-member(2)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.OleFormat): void;
```

Parameters
- properties: [Word.OleFormat](/en-us/javascript/api/word/word.oleformat)

Returns
- void

### toJSON()
Id: word-word-oleformat-tojson-member(1)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.OleFormat` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.OleFormatData`) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.OleFormatData;
```

Returns
- [Word.Interfaces.OleFormatData](/en-us/javascript/api/word/word.interfaces.oleformatdata)

### track()
Id: word-word-oleformat-track-member(1)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.OleFormat;
```

Returns
- [Word.OleFormat](/en-us/javascript/api/word/word.oleformat)

### untrack()
Id: word-word-oleformat-untrack-member(1)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

```typescript
untrack(): Word.OleFormat;
```

Returns
- [Word.OleFormat](/en-us/javascript/api/word/word.oleformat)