# Word.Page class

Represents a page in the document. Page objects manage the page layout and content.

Package: https://learn.microsoft.com/en-us/javascript/api/word

Extends: https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientobject

## Remarks

[ API set: WordApiDesktop 1.2 ]

### Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/35-ranges/get-pages.yaml

await Word.run(async (context) => {
  // Gets pages of the selection.
  const pages: Word.PageCollection = context.document.getSelection().pages;
  pages.load();
  await context.sync();

  // Log info for pages included in selection.
  console.log(pages);
  const pagesIndexes = [];
  const pagesText = [];
  for (let i = 0; i < pages.items.length; i++) {
    const page = pages.items[i];
    page.load('index');
    pagesIndexes.push(page);

    const range = page.getRange();
    range.load('text');
    pagesText.push(range);
  }

  await context.sync();

  for (let i = 0; i < pagesIndexes.length; i++) {
    console.log(`Index info for page ${i + 1} in the selection: ${pagesIndexes[i].index}`);
    console.log("Text of that page in the selection:", pagesText[i].text);
  }
});
```

## Properties

- breaks  
  Gets a BreakCollection object that represents the breaks on the page.

- context  
  The request context associated with the object. This connects the add-in's process to the Office host application's process.

- height  
  Gets the height, in points, of the paper defined in the Page Setup dialog box.

- index  
  Gets the index of the page. The page index is 1-based and independent of the user's custom page numbering.

- width  
  Gets the width, in points, of the paper defined in the Page Setup dialog box.

## Methods

- getNext()  
  Gets the next page in the pane. Throws an ItemNotFound error if this page is the last one.

- getNextOrNullObject()  
  Gets the next page. If this page is the last one, then this method will return an object with its isNullObject property set to true. For further information, see *OrNullObject methods and properties.

- getRange(rangeLocation)  
  Gets the whole page, or the starting or ending point of the page, as a range.

- load(options)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

- load(propertyNames)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

- load(propertyNamesAndPaths)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

- toJSON()  
  Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.Page object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.PageData) that contains shallow copies of any loaded child properties from the original object.

- track()  
  Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

- untrack()  
  Release the memory associated with this object, if it has previously been tracked. This call is 