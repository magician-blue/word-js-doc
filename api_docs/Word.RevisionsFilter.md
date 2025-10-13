# Word.RevisionsFilter class

Package: [word](/en-us/javascript/api/word)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents the current settings related to the display of reviewers' comments and revision marks in the document.

Extends: [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

## Properties

- [context](#context) — The request context associated with the object. This connects the add-in's process to the Office host application's process.
- [markup](#markup) — Specifies a RevisionsMarkup value that represents the extent of reviewer markup displayed in the document.
- [reviewers](#reviewers) — Gets the ReviewerCollection object that represents the collection of reviewers of one or more documents.
- [view](#view) — Specifies a RevisionsView value that represents globally whether Word displays the original version of the document or the final version, which might have revisions and formatting changes applied.

## Methods

- [load(options)](#loadoptions) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- [load(propertyNames)](#loadpropertynames) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- [load(propertyNamesAndPaths)](#loadpropertynamesandpaths) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- [set(properties, options)](#setproperties-options) — Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
- [set(properties)](#setproperties) — Sets multiple properties on the object at the same time, based on an existing loaded object.
- [toggleShowAllReviewers()](#toggleshowallreviewers) — Shows or hides all revisions in the document that contain comments and tracked changes.
- [toJSON()](#tojson) — Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify().
- [track()](#track) — Track the object for automatic adjustment based on surrounding changes in the document.
- [untrack()](#untrack) — Release the memory associated with this object, if it has previously been tracked.

## Property Details

### context

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value: [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### markup

Specifies a RevisionsMarkup value that represents the extent of reviewer markup displayed in the document.

```typescript
markup: Word.RevisionsMarkup | "None" | "Simple" | "All";
```

Property Value: [Word.RevisionsMarkup](/en-us/javascript/api/word/word.revisionsmarkup) | "None" | "Simple" | "All"

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

### reviewers

Gets the ReviewerCollection object that represents the collection of reviewers of one or more documents.

```typescript
readonly reviewers: Word.ReviewerCollection;
```

Property Value: [Word.ReviewerCollection](/en-us/javascript/api/word/word.reviewercollection)

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

### view

Specifies a RevisionsView value that represents globally whether Word displays the original version of the document or the final version, which might have revisions and formatting changes applied.

```typescript
view: Word.RevisionsView | "Final" | "Original";
```

Property Value: [Word.RevisionsView](/en-us/javascript/api/word/word.revisionsview) | "Final" | "Original"

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

## Method Details

### load(options)

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(options?: Word.Interfaces.RevisionsFilterLoadOptions): Word.RevisionsFilter;
```

Parameters:
- options: [Word.Interfaces.RevisionsFilterLoadOptions](/en-us/javascript/api/word/word.interfaces.revisionsfilterloadoptions)  
  Provides options for which properties of the object to load.

Returns: [Word.RevisionsFilter](/en-us/javascript/api/word/word.revisionsfilter)

### load(propertyNames)

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.RevisionsFilter;
```

Parameters:
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns: [Word.RevisionsFilter](/en-us/javascript/api/word/word.revisionsfilter)

### load(propertyNamesAndPaths)

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.RevisionsFilter;
```

Parameters:
- propertyNamesAndPaths:  
  {
  select?: string;
  expand?: string;
  }  
  propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

Returns: [Word.RevisionsFilter](/en-us/javascript/api/word/word.revisionsfilter)

### set(properties, options)

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.RevisionsFilterUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

Parameters:
- properties: [Word.Interfaces.RevisionsFilterUpdateData](/en-us/javascript/api/word/word.interfaces.revisionsfilterupdatedata)  
  A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
- options: [OfficeExtension.UpdateOptions](/en-us/javascript/api/office/officeextension.updateoptions)  
  Provides an option to suppress errors if the properties object tries to set any read-only properties.

Returns: void

### set(properties)

Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.RevisionsFilter): void;
```

Parameters:
- properties: [Word.RevisionsFilter](/en-us/javascript/api/word/word.revisionsfilter)

Returns: void

### toggleShowAllReviewers()

Shows or hides all revisions in the document that contain comments and tracked changes.

```typescript
toggleShowAllReviewers(): void;
```

Returns: void

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

### toJSON()

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.RevisionsFilter object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.RevisionsFilterData) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.RevisionsFilterData;
```

Returns: [Word.Interfaces.RevisionsFilterData](/en-us/javascript/api/word/word.interfaces.revisionsfilterdata)

### track()

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.RevisionsFilter;
```

Returns: [Word.RevisionsFilter](/en-us/javascript/api/word/word.revisionsfilter)

### untrack()

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.RevisionsFilter;
```

Returns: [Word.RevisionsFilter](/en-us/javascript/api/word/word.revisionsfilter)