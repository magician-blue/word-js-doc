# Word.CritiqueAnnotation class

Package: https://learn.microsoft.com/en-us/javascript/api/word

Represents an annotation wrapper around critique displayed in the document.

Extends
- https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientobject

## Remarks

[API set: WordApi 1.7]

Examples
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

## Properties

- context  
  The request context associated with the object. This connects the add-in's process to the Office host application's process.

- critique  
  Gets the critique that was passed when the annotation was inserted.

- range  
  Gets the range of text that is annotated.

## Methods

- accept()  
  Accepts the critique. This will change the annotation state to `accepted`.

- load(options)  
  Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

- load(propertyNames)  
  Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

- load(propertyNamesAndPaths)  
  Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

- reject()  
  Rejects the critique. This will change the annotation state to `rejected`.

- toJSON()  
  Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.CritiqueAnnotation` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.CritiqueAnnotationData`) that contains shallow copies of any loaded child properties from the original object.

- track()  
  Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member. If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

- untrack()  
  Release the memory associated with this object, if it has previously been tracked. This call is shorthand for https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

## Property Details

### context

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value
- https://learn.microsoft.com/en-us/javascript/api/word/word.requestcontext

### critique

Gets the critique that was passed when the annotation was inserted.

```typescript
readonly critique: Word.Critique;
```

Property Value
- https://learn.microsoft.com/en-us/javascript/api/word/word.critique

Remarks
[API set: WordApi 1.7]

Examples
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

### range

Gets the range of text that is annotated.

```typescript
readonly range: Word.Range;
```

Property Value
- https://learn.microsoft.com/en-us/javascript/api/word/word.range

Remarks
[API set: WordApi 1.7]

## Method Details

### accept()

Accepts the critique. This will change the annotation state to `accepted`.

```typescript
accept(): void;
```

Returns
- void

Remarks
[API set: WordApi 1.7]

Examples
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

### load(options)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(options?: Word.Interfaces.CritiqueAnnotationLoadOptions): Word.CritiqueAnnotation;
```

Parameters
- options: https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.critiqueannotationloadoptions  
  Provides options for which properties of the object to load.

Returns
- https://learn.microsoft.com/en-us/javascript/api/word/word.critiqueannotation

### load(propertyNames)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.CritiqueAnnotation;
```

Parameters
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns
- https://learn.microsoft.com/en-us/javascript/api/word/word.critiqueannotation

### load(propertyNamesAndPaths)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
  select?: string;
  expand?: string;
}): Word.CritiqueAnnotation;
```

Parameters
- propertyNamesAndPaths:  
  `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

Returns
- https://learn.microsoft.com/en-us/javascript/api/word/word.critiqueannotation

### reject()

Rejects the critique. This will change the annotation state to `rejected`.

```typescript
reject(): void;
```

Returns
- void

Remarks
[API set: WordApi 1.7]

Examples
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

### toJSON()

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.CritiqueAnnotation` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.CritiqueAnnotationData`) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.CritiqueAnnotationData;
```

Returns
- https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.critiqueannotationdata

### track()

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member. If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.CritiqueAnnotation;
```

Returns
- https://learn.microsoft.com/en-us/javascript/api/word/word.critiqueannotation

### untrack()

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

```typescript
untrack(): Word.CritiqueAnnotation;
```

Returns
- https://learn.microsoft.com/en-us/javascript/api/word/word.critiqueannotation