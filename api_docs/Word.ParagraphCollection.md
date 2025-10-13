# Word.ParagraphCollection class

Package: [word](/en-us/javascript/api/word)

Contains a collection of [Word.Paragraph](/en-us/javascript/api/word/word.paragraph) objects.

Extends: [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks

[ API set: WordApi 1.1 ]

#### Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/15-images/insert-and-get-pictures.yaml

// Inserts an image anchored to the last paragraph.
await Word.run(async (context) => {
  context.document.body.paragraphs
    .getLast()
    .insertParagraph("", "After")
    .insertInlinePictureFromBase64(base64Image, "End");

  await context.sync();
});
```

## Properties

- context  
  The request context associated with the object. This connects the add-in's process to the Office host application's process.

- items  
  Gets the loaded child items in this collection.

## Methods

- add(range)  
  Returns a `Paragraph` object that represents a new, blank paragraph added to the document.

- closeUp()  
  Removes any spacing before the specified paragraphs.

- decreaseSpacing()  
  Decreases the spacing before and after paragraphs in six-point increments.

- getFirst()  
  Gets the first paragraph in this collection. Throws an `ItemNotFound` error if the collection is empty.

- getFirstOrNullObject()  
  Gets the first paragraph in this collection. If the collection is empty, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

- getLast()  
  Gets the last paragraph in this collection. Throws an `ItemNotFound` error if the collection is empty.

- getLastOrNullObject()  
  Gets the last paragraph in this collection. If the collection is empty, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

- increaseSpacing()  
  Increases the spacing before and after paragraphs in six-point increments.

- indent()  
  Indents the paragraphs by one level.

- indentCharacterWidth(count)  
  Indents the paragraphs in the collection by the specified number of characters.

- indentFirstLineCharacterWidth(count)  
  Indents the first line of the paragraphs in the collection by the specified number of characters.

- load(options)  
  Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

- load(propertyNames)  
  Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

- load(propertyNamesAndPaths)  
  Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

- openOrCloseUp()  
  Toggles spacing before paragraphs.

- openUp()  
  Sets spacing before the specified paragraphs to 12 points.

- outdent()  
  Removes one level of indent for the paragraphs.

- outlineDemote()  
  Applies the next heading level style (Heading 1 through Heading 8) to the specified paragraphs.

- outlineDemoteToBody()  
  Demotes the specified paragraphs to body text by applying the Normal style.

- outlinePromote()  
  Applies the previous heading level style (Heading 1 through Heading 8) to the paragraphs in the collection.

- space1()  
  Sets the specified paragraphs to single spacing.

- space1Pt5()  
  Sets the specified paragraphs to 1.5-line spacing.

- space2()  
  Sets the specified paragraphs to double spacing.

- tabHangingIndent(count)  
  Sets a hanging indent to the specified number of tab stops.

- tabIndent(count)  
  Sets the left indent for the specified paragraphs to the specified number of tab stops.

- toJSON()  
  Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`.

- track()  
  Track the object for automatic adjustment based on surrounding changes in the document.

- untrack()  
  Release the memory associated with this object, if it has previously been tracked.

## Property Details

### context

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value  
[Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### items

Gets the loaded child items in this collection.

```typescript
readonly items: Word.Paragraph[];
```

Property Value  
[Word.Paragraph](/en-us/javascript/api/word/word.paragraph)[]

#### Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/20-lists/insert-list.yaml

// This example starts a new list with the second paragraph.
await Word.run(async (context) => {
  const paragraphs: Word.ParagraphCollection = context.document.body.paragraphs;
  paragraphs.load("$none");

  await context.sync();

  // Start new list using the second paragraph.
  const list: Word.List = paragraphs.items[1].startNewList();
  list.load("$none");

  await context.sync();

  // To add new items to the list, use Start or End on the insertLocation parameter.
  list.insertParagraph("New list item at the start of the list", "Start");
  const paragraph: Word.Paragraph = list.insertParagraph("New list item at the end of the list (set to list level 5)", "End");

  // Set up list level for the list item.
  paragraph.listItem.level = 4;

  // To add paragraphs outside the list, use Before or After.
  list.insertParagraph("New paragraph goes after (not part of the list)", "After");

  await context.sync();
});
```

## Method Details

### add(range)

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `Paragraph` object that represents a new, blank paragraph added to the document.

```typescript
add(range: Word.Range): Word.Paragraph;
```

Parameters
- range: [Word.Range](/en-us/javascript/api/word/word.range)  
  The range before which you want the new paragraph to be added. The new paragraph doesn't replace the range.

Returns  
[Word.Paragraph](/en-us/javascript/api/word/word.paragraph)

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### closeUp()

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Removes any spacing before the specified paragraphs.

```typescript
closeUp(): void;
```

Returns  
void

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### decreaseSpacing()

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Decreases the spacing before and after paragraphs in six-point increments.

```typescript
decreaseSpacing(): void;
```

Returns  
void

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### getFirst()

Gets the first paragraph in this collection. Throws an `ItemNotFound` error if the collection is empty.

```typescript
getFirst(): Word.Paragraph;
```

Returns  
[Word.Paragraph](/en-us/javascript/api/word/word.paragraph)

Remarks  
[ API set: WordApi 1.3 ]

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

---

### getFirstOrNullObject()

Gets the first paragraph in this collection. If the collection is empty, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
getFirstOrNullObject(): Word.Paragraph;
```

Returns  
[Word.Paragraph](/en-us/javascript/api/word/word.paragraph)

Remarks  
[ API set: WordApi 1.3 ]

---

### getLast()

Gets the last paragraph in this collection. Throws an `ItemNotFound` error if the collection is empty.

```typescript
getLast(): Word.Paragraph;
```

Returns  
[Word.Paragraph](/en-us/javascript/api/word/word.paragraph)

Remarks  
[ API set: WordApi 1.3 ]

#### Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/15-images/insert-and-get-pictures.yaml

// Inserts an image anchored to the last paragraph.
await Word.run(async (context) => {
  context.document.body.paragraphs
    .getLast()
    .insertParagraph("", "After")
    .insertInlinePictureFromBase64(base64Image, "End");

  await context.sync();
});
```

---

### getLastOrNullObject()

Gets the last paragraph in this collection. If the collection is empty, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
getLastOrNullObject(): Word.Paragraph;
```

Returns  
[Word.Paragraph](/en-us/javascript/api/word/word.paragraph)

Remarks  
[ API set: WordApi 1.3 ]

---

### increaseSpacing()

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Increases the spacing before and after paragraphs in six-point increments.

```typescript
increaseSpacing(): void;
```

Returns  
void

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### indent()

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Indents the paragraphs by one level.

```typescript
indent(): void;
```

Returns  
void

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### indentCharacterWidth(count)

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Indents the paragraphs in the collection by the specified number of characters.

```typescript
indentCharacterWidth(count: number): void;
```

Parameters
- count: number  
  The number of characters by which the specified paragraphs are to be indented.

Returns  
void

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### indentFirstLineCharacterWidth(count)

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Indents the first line of the paragraphs in the collection by the specified number of characters.

```typescript
indentFirstLineCharacterWidth(count: number): void;
```

Parameters
- count: number  
  The number of characters by which the first line of each specified paragraph is to be indented.

Returns  
void

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### load(options)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(options?: Word.Interfaces.ParagraphCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions): Word.ParagraphCollection;
```

Parameters
- options: [Word.Interfaces.ParagraphCollectionLoadOptions](/en-us/javascript/api/word/word.interfaces.paragraphcollectionloadoptions) & [Word.Interfaces.CollectionLoadOptions](/en-us/javascript/api/word/word.interfaces.collectionloadoptions)  
  Provides options for which properties of the object to load.

Returns  
[Word.ParagraphCollection](/en-us/javascript/api/word/word.paragraphcollection)

#### Examples

```typescript
// This example shows how to get the paragraphs in the Word document
// along with their text and font size properties.
// 
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Create a proxy object for the paragraphs collection.
    const paragraphs = context.document.body.paragraphs;

    // Queue a command to load the text and font properties.
    // It is best practice to always specify the property set. Otherwise, all properties are
    // returned in on the object.
    paragraphs.load('text, font/size');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    await context.sync();

    // Insert code that works with the paragraphs loaded by paragraphs.load().
});
```

---

### load(propertyNames)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.ParagraphCollection;
```

Parameters
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns  
[Word.ParagraphCollection](/en-us/javascript/api/word/word.paragraphcollection)

---

### load(propertyNamesAndPaths)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Word.ParagraphCollection;
```

Parameters
- propertyNamesAndPaths: [OfficeExtension.LoadOption](/en-us/javascript/api/office/officeextension.loadoption)  
  `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

Returns  
[Word.ParagraphCollection](/en-us/javascript/api/word/word.paragraphcollection)

---

### openOrCloseUp()

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Toggles spacing before paragraphs.

```typescript
openOrCloseUp(): void;
```

Returns  
void

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### openUp()

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets spacing before the specified paragraphs to 12 points.

```typescript
openUp(): void;
```

Returns  
void

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### outdent()

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Removes one level of indent for the paragraphs.

```typescript
outdent(): void;
```

Returns  
void

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### outlineDemote()

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Applies the next heading level style (Heading 1 through Heading 8) to the specified paragraphs.

```typescript
outlineDemote(): void;
```

Returns  
void

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### outlineDemoteToBody()

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Demotes the specified paragraphs to body text by applying the Normal style.

```typescript
outlineDemoteToBody(): void;
```

Returns  
void

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### outlinePromote()

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Applies the previous heading level style (Heading 1 through Heading 8) to the paragraphs in the collection.

```typescript
outlinePromote(): void;
```

Returns  
void

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### space1()

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets the specified paragraphs to single spacing.

```typescript
space1(): void;
```

Returns  
void

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### space1Pt5()

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets the specified paragraphs to 1.5-line spacing.

```typescript
space1Pt5(): void;
```

Returns  
void

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### space2()

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets the specified paragraphs to double spacing.

```typescript
space2(): void;
```

Returns  
void

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### tabHangingIndent(count)

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets a hanging indent to the specified number of tab stops.

```typescript
tabHangingIndent(count: number): void;
```

Parameters
- count: number  
  The number of tab stops to indent (if positive) or the number of tab stops to remove from the indent (if negative).

Returns  
void

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### tabIndent(count)

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets the left indent for the specified paragraphs to the specified number of tab stops.

```typescript
tabIndent(count: number): void;
```

Parameters
- count: number  
  The number of tab stops to indent (if positive) or the number of tab stops to remove from the indent (if negative).

Returns  
void

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### toJSON()

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.ParagraphCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.ParagraphCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

```typescript
toJSON(): Word.Interfaces.ParagraphCollectionData;
```

Returns  
[Word.Interfaces.ParagraphCollectionData](/en-us/javascript/api/word/word.interfaces.paragraphcollectiondata)

---

### track()

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.ParagraphCollection;
```

Returns  
[Word.ParagraphCollection](/en-us/javascript/api/word/word.paragraphcollection)

---

### untrack()

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

```typescript
untrack(): Word.ParagraphCollection;
```

Returns  
[Word.ParagraphCollection](/en-us/javascript/api/word/word.paragraphcollection)