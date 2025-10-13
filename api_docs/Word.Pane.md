# Word.Pane class

Represents a window pane. The Pane object is a member of the pane collection. The pane collection includes all the window panes for a single window.

- Package: [word](/en-us/javascript/api/word)
- Extends: [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks

[ API set: WordApiDesktop 1.2 ]

#### Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/35-ranges/get-pages.yaml

await Word.run(async (context) => {
  // Gets the first paragraph of each page.
  console.log("Getting first paragraph of each page...");

  // Get the active window.
  const activeWindow: Word.Window = context.document.activeWindow;
  activeWindow.load();

  // Get the active pane.
  const activePane: Word.Pane = activeWindow.activePane;
  activePane.load();

  // Get all pages.
  const pages: Word.PageCollection = activePane.pages;
  pages.load();

  await context.sync();

  // Get page index and paragraphs of each page.
  const pagesIndexes = [];
  const pagesNumberOfParagraphs = [];
  const pagesFirstParagraphText = [];
  for (let i = 0; i < pages.items.length; i++) {
    const page = pages.items[i];
    page.load('index');
    pagesIndexes.push(page);

    const paragraphs = page.getRange().paragraphs;
    paragraphs.load('items/length');
    pagesNumberOfParagraphs.push(paragraphs);

    const firstParagraph = paragraphs.getFirst();
    firstParagraph.load('text');
    pagesFirstParagraphText.push(firstParagraph);
  }

  await context.sync();

  for (let i = 0; i < pagesIndexes.length; i++) {
    console.log(`Page index: ${pagesIndexes[i].index}`);
    console.log(`Number of paragraphs: ${pagesNumberOfParagraphs[i].items.length}`);
    console.log("First paragraph's text:", pagesFirstParagraphText[i].text);
  }
});
```

## Properties

- [context](#context)  
  The request context associated with the object. This connects the add-in's process to the Office host application's process.
- [pages](#pages)  
  Gets the collection of pages in the pane.
- [pagesEnclosingViewport](#pagesenclosingviewport)  
  Gets the PageCollection shown in the viewport of the pane. If a page is partially visible in the pane, the whole page is returned.

## Methods

- [getNext()](#getnext)  
  Gets the next pane in the window. Throws an ItemNotFound error if this pane is the last one.
- [getNextOrNullObject()](#getnextornullobject)  
  Gets the next pane. If this pane is the last one, then this method will return an object with its isNullObject property set to true. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).
- [load(propertyNames)](#loadpropertynames)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- [load(propertyNamesAndPaths)](#loadpropertynamesandpaths)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- [toJSON()](#tojson)  
  Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.Pane object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.PaneData) that contains shallow copies of any loaded child properties from the original object.
- [track()](#track)  
  Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
- [untrack()](#untrack)  
  Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

## Property Details

### context

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value  
[Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### pages

Gets the collection of pages in the pane.

```typescript
readonly pages: Word.PageCollection;
```

Property Value  
[Word.PageCollection](/en-us/javascript/api/word/word.pagecollection)

Remarks  
[ API set: WordApiDesktop 1.2 ]

#### Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/35-ranges/get-pages.yaml

await Word.run(async (context) => {
  // Gets the first paragraph of each page.
  console.log("Getting first paragraph of each page...");

  // Get the active window.
  const activeWindow: Word.Window = context.document.activeWindow;
  activeWindow.load();

  // Get the active pane.
  const activePane: Word.Pane = activeWindow.activePane;
  activePane.load();

  // Get all pages.
  const pages: Word.PageCollection = activePane.pages;
  pages.load();

  await context.sync();

  // Get page index and paragraphs of each page.
  const pagesIndexes = [];
  const pagesNumberOfParagraphs = [];
  const pagesFirstParagraphText = [];
  for (let i = 0; i < pages.items.length; i++) {
    const page = pages.items[i];
    page.load('index');
    pagesIndexes.push(page);

    const paragraphs = page.getRange().paragraphs;
    paragraphs.load('items/length');
    pagesNumberOfParagraphs.push(paragraphs);

    const firstParagraph = paragraphs.getFirst();
    firstParagraph.load('text');
    pagesFirstParagraphText.push(firstParagraph);
  }

  await context.sync();

  for (let i = 0; i < pagesIndexes.length; i++) {
    console.log(`Page index: ${pagesIndexes[i].index}`);
    console.log(`Number of paragraphs: ${pagesNumberOfParagraphs[i].items.length}`);
    console.log("First paragraph's text:", pagesFirstParagraphText[i].text);
  }
});
```

### pagesEnclosingViewport

Gets the PageCollection shown in the viewport of the pane. If a page is partially visible in the pane, the whole page is returned.

```typescript
readonly pagesEnclosingViewport: Word.PageCollection;
```

Property Value  
[Word.PageCollection](/en-us/javascript/api/word/word.pagecollection)

Remarks  
[ API set: WordApiDesktop 1.2 ]

#### Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/35-ranges/get-pages.yaml

await Word.run(async (context) => {
  // Gets the pages enclosing the viewport.

  // Get the active window.
  const activeWindow: Word.Window = context.document.activeWindow;
  activeWindow.load();

  // Get the active pane.
  const activePane: Word.Pane = activeWindow.activePane;
  activePane.load();

  // Get pages enclosing the viewport.
  const pages: Word.PageCollection = activePane.pagesEnclosingViewport;
  pages.load();

  await context.sync();

  // Log the number of pages.
  const pageCount = pages.items.length;
  console.log(`Number of pages enclosing the viewport: ${pageCount}`);

  // Log index info of these pages.
  const pagesIndexes = [];
  for (let i = 0; i < pageCount; i++) {
    const page = pages.items[i];
    page.load('index');
    pagesIndexes.push(page);
  }

  await context.sync();

  for (let i = 0; i < pagesIndexes.length; i++) {
    console.log(`Page index: ${pagesIndexes[i].index}`);
  }
});
```

## Method Details

### getNext()

Gets the next pane in the window. Throws an ItemNotFound error if this pane is the last one.

```typescript
getNext(): Word.Pane;
```

Returns  
[Word.Pane](/en-us/javascript/api/word/word.pane)

Remarks  
[ API set: WordApiDesktop 1.2 ]

### getNextOrNullObject()

Gets the next pane. If this pane is the last one, then this method will return an object with its isNullObject property set to true. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
getNextOrNullObject(): Word.Pane;
```

Returns  
[Word.Pane](/en-us/javascript/api/word/word.pane)

Remarks  
[ API set: WordApiDesktop 1.2 ]

### load(propertyNames)

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.Pane;
```

Parameters  
- propertyNames  
  string | string[]

A comma-delimited string or an array of strings that specify the properties to load.

Returns  
[Word.Pane](/en-us/javascript/api/word/word.pane)

### load(propertyNamesAndPaths)

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.Pane;
```

Parameters  
- propertyNamesAndPaths  
  {
  select?: string;
  expand?: string;
  }

propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

Returns  
[Word.Pane](/en-us/javascript/api/word/word.pane)

### toJSON()

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.Pane object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.PaneData) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.PaneData;
```

Returns  
[Word.Interfaces.PaneData](/en-us/javascript/api/word/word.interfaces.panedata)

### track()

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.Pane;
```

Returns  
[Word.Pane](/en-us/javascript/api/word/word.pane)

### untrack()

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.Pane;
```

Returns  
[Word.Pane](/en-us/javascript/api/word/word.pane)