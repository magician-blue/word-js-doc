# Word.TrackedChange class

Package: [word](/en-us/javascript/api/word)

Represents a tracked change in a Word document.

Extends
- [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks

[API set: WordApi 1.6]

#### Examples
```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-tracked-changes.yaml

// Gets the next (second) tracked change.
await Word.run(async (context) => {
  const body: Word.Body = context.document.body;
  const trackedChanges: Word.TrackedChangeCollection = body.getTrackedChanges();
  await context.sync();

  const trackedChange: Word.TrackedChange = trackedChanges.getFirst();
  await context.sync();

  const nextTrackedChange: Word.TrackedChange = trackedChange.getNext();
  await context.sync();

  nextTrackedChange.load(["author", "date", "text", "type"]);
  await context.sync();

  console.log(nextTrackedChange);
});
```

## Properties

- author: Gets the author of the tracked change.
- context: The request context associated with the object. This connects the add-in's process to the Office host application's process.
- date: Gets the date of the tracked change.
- text: Gets the text of the tracked change.
- type: Gets the type of the tracked change.

## Methods

- accept(): Accepts the tracked change.
- getNext(): Gets the next tracked change. Throws an ItemNotFound error if this tracked change is the last one.
- getNextOrNullObject(): Gets the next tracked change. If this tracked change is the last one, then this method will return an object with its isNullObject property set to true. For further information, see [OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).
- getRange(rangeLocation): Gets the range of the tracked change.
- load(options): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNames): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNamesAndPaths): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- reject(): Rejects the tracked change.
- toJSON(): Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.TrackedChange object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.TrackedChangeData) that contains shallow copies of any loaded child properties from the original object.
- track(): Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
- untrack(): Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

## Property Details

### author

Gets the author of the tracked change.

```typescript
readonly author: string;
```

Property Value
- string

Remarks  
[API set: WordApi 1.6]

### context

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value
- [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### date

Gets the date of the tracked change.

```typescript
readonly date: Date;
```

Property Value
- Date

Remarks  
[API set: WordApi 1.6]

### text

Gets the text of the tracked change.

```typescript
readonly text: string;
```

Property Value
- string

Remarks  
[API set: WordApi 1.6]

### type

Gets the type of the tracked change.

```typescript
readonly type: Word.TrackedChangeType | "None" | "Added" | "Deleted" | "Formatted";
```

Property Value
- [Word.TrackedChangeType](/en-us/javascript/api/word/word.trackedchangetype) | "None" | "Added" | "Deleted" | "Formatted"

Remarks  
[API set: WordApi 1.6]

## Method Details

### accept()

Accepts the tracked change.

```typescript
accept(): void;
```

Returns
- void

Remarks  
[API set: WordApi 1.6]

#### Examples
```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-tracked-changes.yaml

// Accepts the first tracked change.
await Word.run(async (context) => {
  const body: Word.Body = context.document.body;
  const trackedChanges: Word.TrackedChangeCollection = body.getTrackedChanges();
  const trackedChange: Word.TrackedChange = trackedChanges.getFirst();
  trackedChange.load();
  await context.sync();

  console.log("First tracked change:", trackedChange);
  trackedChange.accept();
  console.log("Accepted the first tracked change.");
});
```

### getNext()

Gets the next tracked change. Throws an ItemNotFound error if this tracked change is the last one.

```typescript
getNext(): Word.TrackedChange;
```

Returns
- [Word.TrackedChange](/en-us/javascript/api/word/word.trackedchange)

Remarks  
[API set: WordApi 1.6]

#### Examples
```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-tracked-changes.yaml

// Gets the next (second) tracked change.
await Word.run(async (context) => {
  const body: Word.Body = context.document.body;
  const trackedChanges: Word.TrackedChangeCollection = body.getTrackedChanges();
  await context.sync();

  const trackedChange: Word.TrackedChange = trackedChanges.getFirst();
  await context.sync();

  const nextTrackedChange: Word.TrackedChange = trackedChange.getNext();
  await context.sync();

  nextTrackedChange.load(["author", "date", "text", "type"]);
  await context.sync();

  console.log(nextTrackedChange);
});
```

### getNextOrNullObject()

Gets the next tracked change. If this tracked change is the last one, then this method will return an object with its isNullObject property set to true. For further information, see [OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
getNextOrNullObject(): Word.TrackedChange;
```

Returns
- [Word.TrackedChange](/en-us/javascript/api/word/word.trackedchange)

Remarks  
[API set: WordApi 1.6]

### getRange(rangeLocation)

Gets the range of the tracked change.

```typescript
getRange(rangeLocation?: Word.RangeLocation.whole | Word.RangeLocation.start | Word.RangeLocation.end | "Whole" | "Start" | "End"): Word.Range;
```

Parameters
- rangeLocation  
  [whole](/en-us/javascript/api/word/word.rangelocation#word-word-rangelocation-whole-member) | [start](/en-us/javascript/api/word/word.rangelocation#word-word-rangelocation-start-member) | [end](/en-us/javascript/api/word/word.rangelocation#word-word-rangelocation-end-member) | "Whole" | "Start" | "End"

Returns
- [Word.Range](/en-us/javascript/api/word/word.range)

Remarks  
[API set: WordApi 1.6]

#### Examples
```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-tracked-changes.yaml

// Gets the range of the first tracked change.
await Word.run(async (context) => {
  const body: Word.Body = context.document.body;
  const trackedChanges: Word.TrackedChangeCollection = body.getTrackedChanges();
  const trackedChange: Word.TrackedChange = trackedChanges.getFirst();
  await context.sync();

  const range: Word.Range = trackedChange.getRange();
  range.load();
  await context.sync();

  console.log("range.text: " + range.text);
});
```

### load(options)

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(options?: Word.Interfaces.TrackedChangeLoadOptions): Word.TrackedChange;
```

Parameters
- options  
  [Word.Interfaces.TrackedChangeLoadOptions](/en-us/javascript/api/word/word.interfaces.trackedchangeloadoptions)  
  Provides options for which properties of the object to load.

Returns
- [Word.TrackedChange](/en-us/javascript/api/word/word.trackedchange)

### load(propertyNames)

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.TrackedChange;
```

Parameters
- propertyNames  
  string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns
- [Word.TrackedChange](/en-us/javascript/api/word/word.trackedchange)

### load(propertyNamesAndPaths)

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.TrackedChange;
```

Parameters
- propertyNamesAndPaths  
  {
  select?: string;
  expand?: string;
  }  
  propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

Returns
- [Word.TrackedChange](/en-us/javascript/api/word/word.trackedchange)

### reject()

Rejects the tracked change.

```typescript
reject(): void;
```

Returns
- void

Remarks  
[API set: WordApi 1.6]

#### Examples
```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-tracked-changes.yaml

// Rejects the first tracked change.
await Word.run(async (context) => {
  const body: Word.Body = context.document.body;
  const trackedChanges: Word.TrackedChangeCollection = body.getTrackedChanges();
  const trackedChange: Word.TrackedChange = trackedChanges.getFirst();
  trackedChange.load();
  await context.sync();

  console.log("First tracked change:", trackedChange);
  trackedChange.reject();
  console.log("Rejected the first tracked change.");
});
```

### toJSON()

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.TrackedChange object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.TrackedChangeData) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.TrackedChangeData;
```

Returns
- [Word.Interfaces.TrackedChangeData](/en-us/javascript/api/word/word.interfaces.trackedchangedata)

### track()

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.TrackedChange;
```

Returns
- [Word.TrackedChange](/en-us/javascript/api/word/word.trackedchange)

### untrack()

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.TrackedChange;
```

Returns
- [Word.TrackedChange](/en-us/javascript/api/word/word.trackedchange)