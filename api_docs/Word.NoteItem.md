# Word.NoteItem class

Package: [word](/en-us/javascript/api/word)

Represents a footnote or endnote.

Extends
- [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks
[ API set: WordApi 1.5 ]

#### Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-footnotes.yaml

// Gets the text of the referenced footnote.
await Word.run(async (context) => {
  const footnotes: Word.NoteItemCollection = context.document.body.footnotes;
  footnotes.load("items/body");
  await context.sync();

  const referenceNumber = (document.getElementById("input-reference") as HTMLInputElement).value;
  const mark = (referenceNumber as number) - 1;
  const footnoteBody: Word.Range = footnotes.items[mark].body.getRange();
  footnoteBody.load("text");
  await context.sync();

  console.log(`Text of footnote ${referenceNumber}: ${footnoteBody.text}`);
});
```

## Properties
- body  
  Represents the body object of the note item. It's the portion of the text within the footnote or endnote.
- context  
  The request context associated with the object. This connects the add-in's process to the Office host application's process.
- reference  
  Represents a footnote or endnote reference in the main document.
- type  
  Represents the note item type: footnote or endnote.

## Methods
- delete()  
  Deletes the note item.
- getNext()  
  Gets the next note item of the same type. Throws an `ItemNotFound` error if this note item is the last one.
- getNextOrNullObject()  
  Gets the next note item of the same type. If this note item is the last one, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).
- load(options)  
  Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- load(propertyNames)  
  Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- load(propertyNamesAndPaths)  
  Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- set(properties, options)  
  Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
- set(properties)  
  Sets multiple properties on the object at the same time, based on an existing loaded object.
- toJSON()  
  Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.NoteItem` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.NoteItemData`) that contains shallow copies of any loaded child properties from the original object.
- track()  
  Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
- untrack()  
  Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

## Property Details

### body
Represents the body object of the note item. It's the portion of the text within the footnote or endnote.

```typescript
readonly body: Word.Body;
```

Property value
- [Word.Body](/en-us/javascript/api/word/word.body)

Remarks  
[ API set: WordApi 1.5 ]

#### Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-footnotes.yaml

// Gets the text of the referenced footnote.
await Word.run(async (context) => {
  const footnotes: Word.NoteItemCollection = context.document.body.footnotes;
  footnotes.load("items/body");
  await context.sync();

  const referenceNumber = (document.getElementById("input-reference") as HTMLInputElement).value;
  const mark = (referenceNumber as number) - 1;
  const footnoteBody: Word.Range = footnotes.items[mark].body.getRange();
  footnoteBody.load("text");
  await context.sync();

  console.log(`Text of footnote ${referenceNumber}: ${footnoteBody.text}`);
});
```

### context
The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property value
- [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### reference
Represents a footnote or endnote reference in the main document.

```typescript
readonly reference: Word.Range;
```

Property value
- [Word.Range](/en-us/javascript/api/word/word.range)

Remarks  
[ API set: WordApi 1.5 ]

#### Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-footnotes.yaml

// Selects the footnote's reference mark in the document body.
await Word.run(async (context) => {
  const footnotes: Word.NoteItemCollection = context.document.body.footnotes;
  footnotes.load("items/reference");
  await context.sync();

  const referenceNumber = (document.getElementById("input-reference") as HTMLInputElement).value;
  const mark = (referenceNumber as number) - 1;
  const item: Word.NoteItem = footnotes.items[mark];
  const reference: Word.Range = item.reference;
  reference.select();
  await context.sync();

  console.log(`Reference ${referenceNumber} is selected.`);
});
```

### type
Represents the note item type: footnote or endnote.

```typescript
readonly type: Word.NoteItemType | "Footnote" | "Endnote";
```

Property value
- [Word.NoteItemType](/en-us/javascript/api/word/word.noteitemtype) | "Footnote" | "Endnote"

Remarks  
[ API set: WordApi 1.5 ]

#### Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-footnotes.yaml

// Gets the referenced note's item type and body type, which are both "Footnote".
await Word.run(async (context) => {
  const footnotes: Word.NoteItemCollection = context.document.body.footnotes;
  footnotes.load("items");
  await context.sync();

  const referenceNumber = (document.getElementById("input-reference") as HTMLInputElement).value;
  const mark = (referenceNumber as number) - 1;
  const item: Word.NoteItem = footnotes.items[mark];
  console.log(`Note type of footnote ${referenceNumber}: ${item.type}`);

  item.body.load("type");
  await context.sync();

  console.log(`Body type of note: ${item.body.type}`);
});
```

## Method Details

### delete()
Deletes the note item.

```typescript
delete(): void;
```

Returns
- void

Remarks  
[ API set: WordApi 1.5 ]

#### Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-footnotes.yaml

// Deletes this referenced footnote.
await Word.run(async (context) => {
  const footnotes: Word.NoteItemCollection = context.document.body.footnotes;
  footnotes.load("items");
  await context.sync();

  const referenceNumber = (document.getElementById("input-reference") as HTMLInputElement).value;
  const mark = (referenceNumber as number) - 1;
  footnotes.items[mark].delete();
  await context.sync();

  console.log("Footnote deleted.");
});
```

### getNext()
Gets the next note item of the same type. Throws an `ItemNotFound` error if this note item is the last one.

```typescript
getNext(): Word.NoteItem;
```

Returns
- [Word.NoteItem](/en-us/javascript/api/word/word.noteitem)

Remarks  
[ API set: WordApi 1.5 ]

#### Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-footnotes.yaml

// Selects the next footnote in the document body.
await Word.run(async (context) => {
  const footnotes: Word.NoteItemCollection = context.document.body.footnotes;
  footnotes.load("items/reference");
  await context.sync();

  const referenceNumber = (document.getElementById("input-reference") as HTMLInputElement).value;
  const mark = (referenceNumber as number) - 1;
  const reference: Word.Range = footnotes.items[mark].getNext().reference;
  reference.select();
  console.log("Selected is the next footnote: " + (mark + 2));
});
```

### getNextOrNullObject()
Gets the next note item of the same type. If this note item is the last one, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
getNextOrNullObject(): Word.NoteItem;
```

Returns
- [Word.NoteItem](/en-us/javascript/api/word/word.noteitem)

Remarks  
[ API set: WordApi 1.5 ]

### load(options)
Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(options?: Word.Interfaces.NoteItemLoadOptions): Word.NoteItem;
```

Parameters
- options: [Word.Interfaces.NoteItemLoadOptions](/en-us/javascript/api/word/word.interfaces.noteitemloadoptions)  
  Provides options for which properties of the object to load.

Returns
- [Word.NoteItem](/en-us/javascript/api/word/word.noteitem)

### load(propertyNames)
Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.NoteItem;
```

Parameters
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns
- [Word.NoteItem](/en-us/javascript/api/word/word.noteitem)

### load(propertyNamesAndPaths)
Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.NoteItem;
```

Parameters
- propertyNamesAndPaths:  
  {
  select?: string;
  expand?: string;
  }  
  `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

Returns
- [Word.NoteItem](/en-us/javascript/api/word/word.noteitem)

### set(properties, options)
Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.NoteItemUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

Parameters
- properties: [Word.Interfaces.NoteItemUpdateData](/en-us/javascript/api/word/word.interfaces.noteitemupdatedata)  
  A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
- options: [OfficeExtension.UpdateOptions](/en-us/javascript/api/office/officeextension.updateoptions)  
  Provides an option to suppress errors if the properties object tries to set any read-only properties.

Returns
- void

### set(properties)
Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.NoteItem): void;
```

Parameters
- properties: [Word.NoteItem](/en-us/javascript/api/word/word.noteitem)

Returns
- void

### toJSON()
Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.NoteItem` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.NoteItemData`) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.NoteItemData;
```

Returns
- [Word.Interfaces.NoteItemData](/en-us/javascript/api/word/word.interfaces.noteitemdata)

### track()
Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.NoteItem;
```

Returns
- [Word.NoteItem](/en-us/javascript/api/word/word.noteitem)

### untrack()
Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

```typescript
untrack(): Word.NoteItem;
```

Returns
- [Word.NoteItem](/en-us/javascript/api/word/word.noteitem)