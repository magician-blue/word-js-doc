# Word.StyleCollection class

Package: [word](https://learn.microsoft.com/en-us/javascript/api/word)

Contains a collection of [Word.Style](https://learn.microsoft.com/en-us/javascript/api/word/word.style) objects.

Extends: [OfficeExtension.ClientObject](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientobject)

## Remarks

[API set: WordApi 1.5](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-styles.yaml

// Gets the number of available styles stored with the document.
await Word.run(async (context) => {
  const styles: Word.StyleCollection = context.document.getStyles();
  const count = styles.getCount();
  await context.sync();

  console.log(`Number of styles: ${count.value}`);
});
```

## Properties

- context  
  The request context associated with the object. This connects the add-in's process to the Office host application's process.

- items  
  Gets the loaded child items in this collection.

## Methods

- getByName(name)  
  Get the style object by its name.

- getByNameOrNullObject(name)  
  If the corresponding style doesn't exist, then this method returns an object with its isNullObject property set to true.

- getCount()  
  Gets the number of the styles in the collection.

- getItem(index)  
  Gets a style object by its index in the collection.

- load(options)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

- load(propertyNames)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

- load(propertyNamesAndPaths)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

- toJSON()  
  Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.StyleCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.StyleCollectionData) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

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

#### Property Value
[Word.RequestContext](https://learn.microsoft.com/en-us/javascript/api/word/word.requestcontext)

---

### items

Gets the loaded child items in this collection.

```typescript
readonly items: Word.Style[];
```

#### Property Value
[Word.Style](https://learn.microsoft.com/en-us/javascript/api/word/word.style)[]

## Method Details

### getByName(name)

Get the style object by its name.

```typescript
getByName(name: string): Word.Style;
```

#### Parameters
- name: string  
  Required. The style name.

#### Returns
[Word.Style](https://learn.microsoft.com/en-us/javascript/api/word/word.style)

#### Remarks
[API set: WordApi 1.5](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### getByNameOrNullObject(name)

If the corresponding style doesn't exist, then this method returns an object with its isNullObject property set to true.

```typescript
getByNameOrNullObject(name: string): Word.Style;
```

#### Parameters
- name: string  
  Required. The style name.

#### Returns
[Word.Style](https://learn.microsoft.com/en-us/javascript/api/word/word.style)

#### Remarks
[API set: WordApi 1.5](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-styles.yaml

// Adds a new style.
await Word.run(async (context) => {
  const newStyleName = (document.getElementById("new-style-name") as HTMLInputElement).value;
  if (newStyleName == "") {
    console.warn("Enter a style name to add.");
    return;
  }

  const style: Word.Style = context.document.getStyles().getByNameOrNullObject(newStyleName);
  style.load();
  await context.sync();

  if (!style.isNullObject) {
    console.warn(
      `There's an existing style with the same name '${newStyleName}'! Please provide another style name.`
    );
    return;
  }

  const newStyleType = ((document.getElementById("new-style-type") as HTMLSelectElement).value as unknown) as Word.StyleType;
  context.document.addStyle(newStyleName, newStyleType);
  await context.sync();

  console.log(newStyleName + " has been added to the style list.");
});
```

---

### getCount()

Gets the number of the styles in the collection.

```typescript
getCount(): OfficeExtension.ClientResult<number>;
```

#### Returns
[OfficeExtension.ClientResult](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientresult)<number>

#### Remarks
[API set: WordApi 1.5](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-styles.yaml

// Gets the number of available styles stored with the document.
await Word.run(async (context) => {
  const styles: Word.StyleCollection = context.document.getStyles();
  const count = styles.getCount();
  await context.sync();

  console.log(`Number of styles: ${count.value}`);
});
```

---

### getItem(index)

Gets a style object by its index in the collection.

```typescript
getItem(index: number): Word.Style;
```

#### Parameters
- index: number  
  A number that identifies the index location of a style object.

#### Returns
[Word.Style](https://learn.microsoft.com/en-us/javascript/api/word/word.style)

#### Remarks
[API set: WordApi 1.5](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### load(options)

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(options?: Word.Interfaces.StyleCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions): Word.StyleCollection;
```

#### Parameters
- options: [Word.Interfaces.StyleCollectionLoadOptions](https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.stylecollectionloadoptions) & [Word.Interfaces.CollectionLoadOptions](https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.collectionloadoptions)  
  Provides options for which properties of the object to load.

#### Returns
[Word.StyleCollection](https://learn.microsoft.com/en-us/javascript/api/word/word.stylecollection)

---

### load(propertyNames)

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.StyleCollection;
```

#### Parameters
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

#### Returns
[Word.StyleCollection](https://learn.microsoft.com/en-us/javascript/api/word/word.stylecollection)

---

### load(propertyNamesAndPaths)

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Word.StyleCollection;
```

#### Parameters
- propertyNamesAndPaths: [OfficeExtension.LoadOption](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.loadoption)  
  propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

#### Returns
[Word.StyleCollection](https://learn.microsoft.com/en-us/javascript/api/word/word.stylecollection)

---

### toJSON()

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.StyleCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.StyleCollectionData) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

```typescript
toJSON(): Word.Interfaces.StyleCollectionData;
```

#### Returns
[Word.Interfaces.StyleCollectionData](https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.stylecollectiondata)

---

### track()

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.StyleCollection;
```

#### Returns
[Word.StyleCollection](https://learn.microsoft.com/en-us/javascript/api/word/word.stylecollection)

---

### untrack()

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.StyleCollection;
```

#### Returns
[Word.StyleCollection](https://learn.microsoft.com/en-us/javascript/api/word/word.stylecollection)