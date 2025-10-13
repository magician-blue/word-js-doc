# Word.AnnotationCollection class

Package: [word](/en-us/javascript/api/word)

Contains a collection of [Word.Annotation](/en-us/javascript/api/word/word.annotation) objects.

Extends: [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks

[ API set: WordApi 1.7 ]

#### Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-annotations.yaml

// Gets annotations found in the selected paragraph.
await Word.run(async (context) => {
  const paragraph: Word.Paragraph = context.document.getSelection().paragraphs.getFirst();
  const annotations: Word.AnnotationCollection = paragraph.getAnnotations();
  annotations.load("id,state,critiqueAnnotation");

  await context.sync();

  console.log("Annotations found:");

  for (let i = 0; i < annotations.items.length; i++) {
    const annotation: Word.Annotation = annotations.items[i];

    console.log(`ID ${annotation.id} - state '${annotation.state}':`, annotation.critiqueAnnotation.critique);
  }
});
```

## Properties

- context  
  The request context associated with the object. This connects the add-in's process to the Office host application's process.

- items  
  Gets the loaded child items in this collection.

## Methods

- getFirst()  
  Gets the first annotation in this collection. Throws an ItemNotFound error if this collection is empty.

- getFirstOrNullObject()  
  Gets the first annotation in this collection. If this collection is empty, then this method will return an object with its isNullObject property set to true. For further information, see \*OrNullObject methods and properties.

- load(options)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

- load(propertyNames)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

- load(propertyNamesAndPaths)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

- toJSON()  
  Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.AnnotationCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.AnnotationCollectionData) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

- track()  
  Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

- untrack()  
  Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

## Property Details

### context

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value: [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### items

Gets the loaded child items in this collection.

```typescript
readonly items: Word.Annotation[];
```

Property Value: [Word.Annotation](/en-us/javascript/api/word/word.annotation)[]

## Method Details

### getFirst()

Gets the first annotation in this collection. Throws an ItemNotFound error if this collection is empty.

```typescript
getFirst(): Word.Annotation;
```

Returns: [Word.Annotation](/en-us/javascript/api/word/word.annotation)

Remarks  
[ API set: WordApi 1.7 ]

### getFirstOrNullObject()

Gets the first annotation in this collection. If this collection is empty, then this method will return an object with its isNullObject property set to true. For further information, see [\*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
getFirstOrNullObject(): Word.Annotation;
```

Returns: [Word.Annotation](/en-us/javascript/api/word/word.annotation)

Remarks  
[ API set: WordApi 1.7 ]

### load(options)

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(options?: Word.Interfaces.AnnotationCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions): Word.AnnotationCollection;
```

Parameters
- options: [Word.Interfaces.AnnotationCollectionLoadOptions](/en-us/javascript/api/word/word.interfaces.annotationcollectionloadoptions) & [Word.Interfaces.CollectionLoadOptions](/en-us/javascript/api/word/word.interfaces.collectionloadoptions)  
  Provides options for which properties of the object to load.

Returns: [Word.AnnotationCollection](/en-us/javascript/api/word/word.annotationcollection)

### load(propertyNames)

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.AnnotationCollection;
```

Parameters
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns: [Word.AnnotationCollection](/en-us/javascript/api/word/word.annotationcollection)

### load(propertyNamesAndPaths)

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Word.AnnotationCollection;
```

Parameters
- propertyNamesAndPaths: [OfficeExtension.LoadOption](/en-us/javascript/api/office/officeextension.loadoption)  
  propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

Returns: [Word.AnnotationCollection](/en-us/javascript/api/word/word.annotationcollection)

### toJSON()

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.AnnotationCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.AnnotationCollectionData) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

```typescript
toJSON(): Word.Interfaces.AnnotationCollectionData;
```

Returns: [Word.Interfaces.AnnotationCollectionData](/en-us/javascript/api/word/word.interfaces.annotationcollectiondata)

### track()

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.AnnotationCollection;
```

Returns: [Word.AnnotationCollection](/en-us/javascript/api/word/word.annotationcollection)

### untrack()

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.AnnotationCollection;
```

Returns: [Word.AnnotationCollection](/en-us/javascript/api/word/word.annotationcollection)