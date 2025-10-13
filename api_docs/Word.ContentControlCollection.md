# Word.ContentControlCollection class

- Package: [word](https://learn.microsoft.com/en-us/javascript/api/word)

Contains a collection of [Word.ContentControl](https://learn.microsoft.com/en-us/javascript/api/word/word.contentcontrol) objects. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain contents such as images, tables, or paragraphs of formatted text. Currently, only rich text, plain text, checkbox, dropdown list, and combo box content controls are supported.

- Extends: [OfficeExtension.ClientObject](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientobject)

## Remarks

[API set: WordApi 1.1](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### Examples
```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/90-scenarios/doc-assembly.yaml

await Word.run(async (context) => {
    const contentControls: Word.ContentControlCollection = context.document.contentControls.getByTag("customer");
    contentControls.load("text");

    await context.sync();

    for (let i = 0; i < contentControls.items.length; i++) {
      contentControls.items[i].insertText("Fabrikam", "Replace");
    }

    await context.sync();
});
```

## Properties

- context  
  The request context associated with the object. This connects the add-in's process to the Office host application's process.

- items  
  Gets the loaded child items in this collection.

## Methods

- getByChangeTrackingStates(changeTrackingStates)  
  Gets the content controls that have the specified tracking state.

- getById(id)  
  Gets a content control by its identifier. Throws an `ItemNotFound` error if there isn't a content control with the identifier in this collection.

- getByIdOrNullObject(id)  
  Gets a content control by its identifier. If there isn't a content control with the identifier in this collection, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [\*OrNullObject methods and properties](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

- getByTag(tag)  
  Gets the content controls that have the specified tag.

- getByTitle(title)  
  Gets the content controls that have the specified title.

- getByTypes(types)  
  Gets the content controls that have the specified types.

- getFirst()  
  Gets the first content control in this collection. Throws an `ItemNotFound` error if this collection is empty.

- getFirstOrNullObject()  
  Gets the first content control in this collection. If this collection is empty, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [\*OrNullObject methods and properties](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

- getItem(id)  
  Gets a content control by its ID.

- load(options)  
  Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

- load(propertyNames)  
  Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

- load(propertyNamesAndPaths)  
  Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

- toJSON()  
  Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.ContentControlCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.ContentControlCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

- track()  
  Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

- untrack()  
  Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

## Property Details

### context

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value: [Word.RequestContext](https://learn.microsoft.com/en-us/javascript/api/word/word.requestcontext)

### items

Gets the loaded child items in this collection.

```typescript
readonly items: Word.ContentControl[];
```

Property Value: [Word.ContentControl](https://learn.microsoft.com/en-us/javascript/api/word/word.contentcontrol)[]

## Method Details

### getByChangeTrackingStates(changeTrackingStates)

Gets the content controls that have the specified tracking state.

```typescript
getByChangeTrackingStates(changeTrackingStates: Word.ChangeTrackingState[]): Word.ContentControlCollection;
```

Parameters:
- changeTrackingStates: [Word.ChangeTrackingState](https://learn.microsoft.com/en-us/javascript/api/word/word.changetrackingstate)[]  
  Required. An array of content control change tracking states.

Returns: [Word.ContentControlCollection](https://learn.microsoft.com/en-us/javascript/api/word/word.contentcontrolcollection)

Remarks: [API set: WordApi 1.5](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### getById(id)

Gets a content control by its identifier. Throws an `ItemNotFound` error if there isn't a content control with the identifier in this collection.

```typescript
getById(id: number): Word.ContentControl;
```

Parameters:
- id: number  
  Required. A content control identifier.

Returns: [Word.ContentControl](https://learn.microsoft.com/en-us/javascript/api/word/word.contentcontrol)

Remarks: [API set: WordApi 1.1](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples
```typescript
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Create a proxy object for the content control that contains a specific id.
    const contentControl = context.document.contentControls.getById(30086310);

    // Queue a command to load the text property for a content control.
    contentControl.load('text');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    await context.sync();
    console.log('The content control with that Id has been found in this document.');
});
```

---

### getByIdOrNullObject(id)

Gets a content control by its identifier. If there isn't a content control with the identifier in this collection, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [\*OrNullObject methods and properties](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
getByIdOrNullObject(id: number): Word.ContentControl;
```

Parameters:
- id: number  
  Required. A content control identifier.

Returns: [Word.ContentControl](https://learn.microsoft.com/en-us/javascript/api/word/word.contentcontrol)

Remarks: [API set: WordApi 1.3](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples
```typescript
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Create a proxy object for the content control that contains a specific id.
    const contentControl = context.document.contentControls.getByIdOrNullObject(30086310);

    // Queue a command to load the text property for a content control.
    contentControl.load('text');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    await context.sync();
    if (contentControl.isNullObject) {
        console.log('There is no content control with that ID.')
    } else {
        console.log('The content control with that ID has been found in this document.');
    }
});
```

---

### getByTag(tag)

Gets the content controls that have the specified tag.

```typescript
getByTag(tag: string): Word.ContentControlCollection;
```

Parameters:
- tag: string  
  Required. A tag set on a content control.

Returns: [Word.ContentControlCollection](https://learn.microsoft.com/en-us/javascript/api/word/word.contentcontrolcollection)

Remarks: [API set: WordApi 1.1](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples
```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/90-scenarios/doc-assembly.yaml

await Word.run(async (context) => {
    const contentControls: Word.ContentControlCollection = context.document.contentControls.getByTag("customer");
    contentControls.load("text");

    await context.sync();

    for (let i = 0; i < contentControls.items.length; i++) {
      contentControls.items[i].insertText("Fabrikam", "Replace");
    }

    await context.sync();
});
```

---

### getByTitle(title)

Gets the content controls that have the specified title.

```typescript
getByTitle(title: string): Word.ContentControlCollection;
```

Parameters:
- title: string  
  Required. The title of a content control.

Returns: [Word.ContentControlCollection](https://learn.microsoft.com/en-us/javascript/api/word/word.contentcontrolcollection)

Remarks: [API set: WordApi 1.1](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples
```typescript
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Create a proxy object for the content controls collection that contains a specific title.
    const contentControlsWithTitle = context.document.contentControls.getByTitle('Enter Customer Address Here');

    // Queue a command to load the text property for all of content controls with a specific titl