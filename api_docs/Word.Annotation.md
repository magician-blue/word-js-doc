# Word.Annotation class

Package: [word](/en-us/javascript/api/word)

Represents an annotation attached to a paragraph.

Extends
- [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks

[API set: WordApi 1.7]

#### Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-annotations.yaml

// Accepts the first annotation found in the selected paragraph.
await Word.run(async (context) => {
  const paragraph: Word.Paragraph = context.document.getSelection().paragraphs.getFirst();
  const annotations: Word.AnnotationCollection = paragraph.getAnnotations();
  annotations.load("id,state,critiqueAnnotation");

  await context.sync();

  for (let i = 0; i < annotations.items.length; i++) {
    const annotation: Word.Annotation = annotations.items[i];

    if (annotation.state === Word.AnnotationState.created) {
      console.log(`Accepting ID ${annotation.id}...`);
      annotation.critiqueAnnotation.accept();

      await context.sync();
      break;
    }
  }
});
```

## Properties

- [context](#context) — The request context associated with the object. This connects the add-in's process to the Office host application's process.
- [critiqueAnnotation](#critiqueannotation) — Gets the critique annotation object.
- [id](#id) — Gets the unique identifier, which is meant to be used for easier tracking of Annotation objects.
- [state](#state) — Gets the state of the annotation.

## Methods

- [delete()](#delete) — Deletes the annotation.
- [load(options)](#loadoptions) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- [load(propertyNames)](#loadpropertynames) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- [load(propertyNamesAndPaths)](#loadpropertynamesandpaths) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- [toJSON()](#tojson) — Overrides the JavaScript toJSON() method to provide more useful output when an API object is passed to JSON.stringify(). The method returns a plain JavaScript object (typed as Word.Interfaces.AnnotationData) that contains shallow copies of any loaded child properties.
- [track()](#track) — Track the object for automatic adjustment based on surrounding changes in the document.
- [untrack()](#untrack) — Release the memory associated with this object, if it has previously been tracked.

## Property Details

### context

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value
- [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

---

### critiqueAnnotation

Gets the critique annotation object.

```typescript
readonly critiqueAnnotation: Word.CritiqueAnnotation;
```

Property Value
- [Word.CritiqueAnnotation](/en-us/javascript/api/word/word.critiqueannotation)

Remarks  
[API set: WordApi 1.7]

#### Examples
```TypeScript
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

---

### id

Gets the unique identifier, which is meant to be used for easier tracking of Annotation objects.

```typescript
readonly id: string;
```

Property Value
- string

Remarks  
[API set: WordApi 1.7]

#### Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-annotations.yaml

// Accepts the first annotation found in the selected paragraph.
await Word.run(async (context) => {
  const paragraph: Word.Paragraph = context.document.getSelection().paragraphs.getFirst();
  const annotations: Word.AnnotationCollection = paragraph.getAnnotations();
  annotations.load("id,state,critiqueAnnotation");

  await context.sync();

  for (let i = 0; i < annotations.items.length; i++) {
    const annotation: Word.Annotation = annotations.items[i];

    if (annotation.state === Word.AnnotationState.created) {
      console.log(`Accepting ID ${annotation.id}...`);
      annotation.critiqueAnnotation.accept();

      await context.sync();
      break;
    }
  }
});
```

---

### state

Gets the state of the annotation.

```typescript
readonly state: Word.AnnotationState | "Created" | "Accepted" | "Rejected";
```

Property Value
- [Word.AnnotationState](/en-us/javascript/api/word/word.annotationstate) | "Created" | "Accepted" | "Rejected"

Remarks  
[API set: WordApi 1.7]

#### Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-annotations.yaml

// Rejects the last annotation found in the selected paragraph.
await Word.run(async (context) => {
  const paragraph: Word.Paragraph = context.document.getSelection().paragraphs.getFirst();
  const annotations: Word.AnnotationCollection = paragraph.getAnnotations();
  annotations.load("id,state,critiqueAnnotation");

  await context.sync();

  for (let i = annotations.items.length - 1; i >= 0; i--) {
    const annotation: Word.Annotation = annotations.items[i];

    if (annotation.state === Word.AnnotationState.created) {
      console.log(`Rejecting ID ${annotation.id}...`);
      annotation.critiqueAnnotation.reject();

      await context.sync();
      break;
    }
  }
});
```

## Method Details

### delete

Deletes the annotation.

```typescript
delete(): void;
```

Returns
- void

Remarks  
[API set: WordApi 1.7]

#### Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-annotations.yaml

// Deletes all annotations found in the selected paragraph.
await Word.run(async (context) => {
  const paragraph: Word.Paragraph = context.document.getSelection().paragraphs.getFirst();
  const annotations: Word.AnnotationCollection = paragraph.getAnnotations();
  annotations.load("id");

  await context.sync();

  const ids = [];
  for (let i = 0; i < annotations.items.length; i++) {
    const annotation: Word.Annotation = annotations.items[i];

    ids.push(annotation.id);
    annotation.delete();
  }

  await context.sync();

  console.log("Annotations deleted:", ids);
});
```

---

### load(options)

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(options?: Word.Interfaces.AnnotationLoadOptions): Word.Annotation;
```

Parameters
- options: [Word.Interfaces.AnnotationLoadOptions](/en-us/javascript/api/word/word.interfaces.annotationloadoptions)  
  Provides options for which properties of the object to load.

Returns
- [Word.Annotation](/en-us/javascript/api/word/word.annotation)

---

### load(propertyNames)

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.Annotation;
```

Parameters
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns
- [Word.Annotation](/en-us/javascript/api/word/word.annotation)

---

### load(propertyNamesAndPaths)

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
  select?: string;
  expand?: string;
}): Word.Annotation;
```

Parameters
- propertyNamesAndPaths:  
  select is a comma-delimited string that specifies the properties to load, and expand is a comma-delimited string that specifies the navigation properties to load.

Returns
- [Word.Annotation](/en-us/javascript/api/word/word.annotation)

---

### toJSON

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.Annotation object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.AnnotationData) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.AnnotationData;
```

Returns
- [Word.Interfaces.AnnotationData](/en-us/javascript/api/word/word.interfaces.annotationdata)

---

### track

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.Annotation;
```

Returns
- [Word.Annotation](/en-us/javascript/api/word/word.annotation)

---

### untrack

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.Annotation;
```

Returns
- [Word.Annotation](/en-us/javascript/api/word/word.annotation)