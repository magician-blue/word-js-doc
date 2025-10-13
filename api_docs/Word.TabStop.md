# Word.TabStop class

Package: [word](/en-us/javascript/api/word)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents a tab stop in a Word document.

Extends: [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks

[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- alignment  
  Gets a TabAlignment value that represents the alignment for the tab stop.
- context  
  The request context associated with the object. This connects the add-in's process to the Office host application's process.
- customTab  
  Gets whether this tab stop is a custom tab stop.
- leader  
  Gets a TabLeader value that represents the leader for this TabStop object.
- next  
  Gets the next tab stop in the collection.
- position  
  Gets the position of the tab stop relative to the left margin.
- previous  
  Gets the previous tab stop in the collection.

## Methods

- clear()  
  Removes this custom tab stop.
- load(options)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNames)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNamesAndPaths)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- toJSON()  
  Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.TabStop object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.TabStopData) that contains shallow copies of any loaded child properties from the original object.
- track()  
  Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
- untrack()  
  Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

## Property Details

### alignment

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets a TabAlignment value that represents the alignment for the tab stop.

```typescript
readonly alignment: Word.TabAlignment | "Left" | "Center" | "Right" | "Decimal" | "Bar" | "List";
```

Property Value  
[Word.TabAlignment](/en-us/javascript/api/word/word.tabalignment) | "Left" | "Center" | "Right" | "Decimal" | "Bar" | "List"

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### context

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value  
[Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### customTab

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether this tab stop is a custom tab stop.

```typescript
readonly customTab: boolean;
```

Property Value  
boolean

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### leader

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets a TabLeader value that represents the leader for this TabStop object.

```typescript
readonly leader: Word.TabLeader | "Spaces" | "Dots" | "Dashes" | "Lines" | "Heavy" | "MiddleDot";
```

Property Value  
[Word.TabLeader](/en-us/javascript/api/word/word.tableader) | "Spaces" | "Dots" | "Dashes" | "Lines" | "Heavy" | "MiddleDot"

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### next

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the next tab stop in the collection.

```typescript
readonly next: Word.TabStop;
```

Property Value  
[Word.TabStop](/en-us/javascript/api/word/word.tabstop)

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### position

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the position of the tab stop relative to the left margin.

```typescript
readonly position: number;
```

Property Value  
number

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### previous

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the previous tab stop in the collection.

```typescript
readonly previous: Word.TabStop;
```

Property Value  
[Word.TabStop](/en-us/javascript/api/word/word.tabstop)

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Method Details

### clear()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Removes this custom tab stop.

```typescript
clear(): void;
```

Returns  
void

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### load(options)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(options?: Word.Interfaces.TabStopLoadOptions): Word.TabStop;
```

Parameters  
- options: [Word.Interfaces.TabStopLoadOptions](/en-us/javascript/api/word/word.interfaces.tabstoploadoptions)  
  Provides options for which properties of the object to load.

Returns  
[Word.TabStop](/en-us/javascript/api/word/word.tabstop)

### load(propertyNames)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.TabStop;
```

Parameters  
- propertyNames: `string | string[]`  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns  
[Word.TabStop](/en-us/javascript/api/word/word.tabstop)

### load(propertyNamesAndPaths)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
    select?: string;
    expand?: string;
}): Word.TabStop;
```

Parameters  
- propertyNamesAndPaths:  
  {
  select?: string;  
  expand?: string;  
  }  
  propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

Returns  
[Word.TabStop](/en-us/javascript/api/word/word.tabstop)

### toJSON()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.TabStop object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.TabStopData) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.TabStopData;
```

Returns  
[Word.Interfaces.TabStopData](/en-us/javascript/api/word/word.interfaces.tabstopdata)

### track()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.TabStop;
```

Returns  
[Word.TabStop](/en-us/javascript/api/word/word.tabstop)

### untrack()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.TabStop;
```

Returns  
[Word.TabStop](/en-us/javascript/api/word/word.tabstop)