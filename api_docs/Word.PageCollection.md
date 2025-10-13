# Word.PageCollection class

Package: [word](https://learn.microsoft.com/en-us/javascript/api/word)

Represents the collection of page.

Extends: [OfficeExtension.ClientObject](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientobject)

## Remarks
[API set: WordApiDesktop 1.2](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples
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
- context  
  The request context associated with the object. This connects the add-in's process to the Office host application's process.

- items  
  Gets the loaded child items in this collection.

## Methods
- getFirst()  
  Gets the first page in this collection. Throws an ItemNotFound error if this collection is empty.

- getFirstOrNullObject()  
  Gets the first page in this collection. If this collection is empty, then this method will return an object with its isNullObject property set to true. For further information, see [\*OrNullObject methods and properties](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

- load(options)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

- load(propertyNames)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

- load(propertyNamesAndPaths)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

- toJSON()  
  Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.PageCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.PageCollectionData) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

- track()  
  Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

- untrack()  
  Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

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
readonly items: Word.Page[];
```

Property Value: [Word.Page](https://learn.microsoft.com/en-us/javascript/api/word/word.page)[]

## Method Details

### getFirst()
Gets the first page in this collection. Throws an ItemNotFound error if this collection is empty.

```typescript
getFirst(): Word.Page;
```

Returns: [Word.Page](https://learn.microsoft.com/en-us/javascript/api/word/word.page)

Remarks  
[API set: WordApiDesktop 1.2](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### getFirstOrNullObject()
Gets the first page in this collection. If this collection is empty, then this method will return an object with its isNullObject property set to true. For further information, see [\*OrNullObject methods and properties](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
getFirstOrNullObject(): Word.Page;
```

Returns: [Word.Page](https://learn.microsoft.com/en-us/javascript/api/word/word.page)

Remarks  
[API set: WordApiDesktop 1.2](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### load(options)
Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(options?: Word.Interfaces.PageCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions): Word.PageCollection;
```

Parameters
- options: [Word.Interfaces.PageCollectionLoadOptions](https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.pagecollectionloadoptions) & [Word.Interfaces.CollectionLoadOptions](https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.collectionloadoptions)  
  Provides options for which properties of the object to load.

Returns: [Word.PageCollection](https://learn.microsoft.com/en-us/javascript/api/word/word.pagecollection)

### load(propertyNames)
Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.PageCollection;
```

Parameters
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns: [Word.PageCollection](https://learn.microsoft.com/en-us/javascript/api/word/word.pagecollection)

### load(propertyNamesAndPaths)
Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Word.PageCollection;
```

Parameters
- propertyNamesAndPaths: [OfficeExtension.LoadOption](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.loadoption)  
  propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

Returns: [Word.PageCollection](https://learn.microsoft.com/en-us/javascript/api/word/word.pagecollection)

### toJSON()
Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.PageCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.PageCollectionData) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

```typescript
toJSON(): Word.Interfaces.PageCollectionData;
```

Returns: [Word.Interfaces.PageCollectionData](https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.pagecollectiondata)

### track()
Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.PageCollection;
```

Returns: [Word.PageCollection](https://learn.microsoft.com/en-us/javascript/api/word/word.pagecollection)

### untrack()
Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.PageCollection;
```

Returns: [Word.PageCollection](https://learn.microsoft.com/en-us/javascript/api/word/word.pagecollection)