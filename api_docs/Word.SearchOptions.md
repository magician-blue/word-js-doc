# Word.SearchOptions class

- Package: [word](https://learn.microsoft.com/en-us/javascript/api/word)

Specifies the options to be included in a search operation. To learn more about how to use search options in the Word JavaScript APIs, read Use search options to find text in your Word add-in.

Extends: [OfficeExtension.ClientObject](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientobject)

## Remarks

[ API set: WordApi 1.1 ]

### Examples

```typescript
// Search using a wildcard
// Run a batch operation against the Word object model.
await Word.run(async (context) => {
    // Queue a command to search the document with a wildcard
    // for any string of characters that starts with 'to' and ends with 'n'.
    const searchResults = context.document.body.search('to*n', {matchWildcards: true});

    // Queue a command to load the search results and get the font property values.
    searchResults.load('font');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    await context.sync();
    console.log('Found count: ' + searchResults.items.length);

    // Queue a set of commands to change the font for each found item.
    for (let i = 0; i < searchResults.items.length; i++) {
        searchResults.items[i].font.color = 'purple';
        searchResults.items[i].font.highlightColor = 'pink';
        searchResults.items[i].font.bold = true;
    }
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    await context.sync();
}); 
```

## Properties

- `context`: The request context associated with the object. This connects the add-in's process to the Office host application's process.
- `ignorePunct`: Specifies a value that indicates whether to ignore all punctuation characters between words. Corresponds to the Ignore punctuation check box in the Find and Replace dialog box.
- `ignoreSpace`: Specifies a value that indicates whether to ignore all whitespace between words. Corresponds to the Ignore whitespace characters check box in the Find and Replace dialog box.
- `matchCase`: Specifies a value that indicates whether to perform a case sensitive search. Corresponds to the Match case check box in the Find and Replace dialog box.
- `matchPrefix`: Specifies a value that indicates whether to match words that begin with the search string. Corresponds to the Match prefix check box in the Find and Replace dialog box.
- `matchSuffix`: Specifies a value that indicates whether to match words that end with the search string. Corresponds to the Match suffix check box in the Find and Replace dialog box.
- `matchWholeWord`: Specifies a value that indicates whether to find operation only entire words, not text that is part of a larger word. Corresponds to the Find whole words only check box in the Find and Replace dialog box.
- `matchWildcards`: Specifies a value that indicates whether the search will be performed using special search operators. Corresponds to the Use wildcards check box in the Find and Replace dialog box.

## Methods

- `load(options)`: Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- `load(propertyNames)`: Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- `load(propertyNamesAndPaths)`: Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- `newObject(context)`: Create a new instance of the `Word.SearchOptions` object.
- `set(properties, options)`: Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
- `set(properties)`: Sets multiple properties on the object at the same time, based on an existing loaded object.
- `toJSON()`: Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.SearchOptions` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.SearchOptionsData`) that contains shallow copies of any loaded child properties from the original object.

## Property Details

### context

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property value: [Word.RequestContext](https://learn.microsoft.com/en-us/javascript/api/word/word.requestcontext)

---

### ignorePunct

Specifies a value that indicates whether to ignore all punctuation characters between words. Corresponds to the Ignore punctuation check box in the Find and Replace dialog box.

```typescript
ignorePunct: boolean;
```

Property value: boolean

Remarks: [ API set: WordApi 1.1 ]

---

### ignoreSpace

Specifies a value that indicates whether to ignore all whitespace between words. Corresponds to the Ignore whitespace characters check box in the Find and Replace dialog box.

```typescript
ignoreSpace: boolean;
```

Property value: boolean

Remarks: [ API set: WordApi 1.1 ]

---

### matchCase

Specifies a value that indicates whether to perform a case sensitive search. Corresponds to the Match case check box in the Find and Replace dialog box.

```typescript
matchCase: boolean;
```

Property value: boolean

Remarks: [ API set: WordApi 1.1 ]

---

### matchPrefix

Specifies a value that indicates whether to match words that begin with the search string. Corresponds to the Match prefix check box in the Find and Replace dialog box.

```typescript
matchPrefix: boolean;
```

Property value: boolean

Remarks: [ API set: WordApi 1.1 ]

---

### matchSuffix

Specifies a value that indicates whether to match words that end with the search string. Corresponds to the Match suffix check box in the Find and Replace dialog box.

```typescript
matchSuffix: boolean;
```

Property value: boolean

Remarks: [ API set: WordApi 1.1 ]

---

### matchWholeWord

Specifies a value that indicates whether to find operation only entire words, not text that is part of a larger word. Corresponds to the Find whole words only check box in the Find and Replace dialog box.

```typescript
matchWholeWord: boolean;
```

Property value: boolean

Remarks: [ API set: WordApi 1.1 ]

---

### matchWildcards

Specifies a value that indicates whether the search will be performed using special search operators. Corresponds to the Use wildcards check box in the Find and Replace dialog box.

```typescript
matchWildcards: boolean;
```

Property value: boolean

Remarks: [ API set: WordApi 1.1 ]

## Method Details

### load(options)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(options?: Word.Interfaces.SearchOptionsLoadOptions): Word.SearchOptions;
```

Parameters:
- `options`: [Word.Interfaces.SearchOptionsLoadOptions](https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.searchoptionsloadoptions)  
  Provides options for which properties of the object to load.

Returns: [Word.SearchOptions](https://learn.microsoft.com/en-us/javascript/api/word/word.searchoptions)

#### Examples

```typescript
// Ignore punctuation search
// Run a batch operation against the Word object model.
await Word.run(async (context) => {
    // Queue a command to search the document and ignore punctuation.
    const searchResults = context.document.body.search('video you', {ignorePunct: true});

    // Queue a command to load the search results and get the font property values.
    searchResults.load('font');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    await context.sync();
    console.log('Found count: ' + searchResults.items.length);

    // Queue a set of commands to change the font for each found item.
    for (let i = 0; i < searchResults.items.length; i++) {
        searchResults.items[i].font.color = 'purple';
        searchResults.items[i].font.highlightColor = '#FFFF00'; //Yellow
        searchResults.items[i].font.bold = true;
    }
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    await context.sync();
});  
```

```typescript
// Search based on a prefix
// Run a batch operation against the Word object model.
await Word.run(async (context) => {
    // Queue a command to search the document based on a prefix.
    const searchResults = context.document.body.search('vid', {matchPrefix: true});

    // Queue a command to load the search results and get the font property values.
    searchResults.load('font');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    await context.sync();

    // Queue a set of commands to change the font for each found item.
    for (let i = 0; i < searchResults.items.length; i++) {
        searchResults.items[i].font.color = 'purple';
        searchResults.items[i].font.highlightColor = '#FFFF00'; //Yellow
        searchResults.items[i].font.bold = true;
    }
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    await context.sync();
}); 
```

```typescript
// Search based on a suffix
// Run a batch operation against the Word object model.
await Word.run(async (context) => {
    // Queue a command to search the document for any string of characters after 'ly'.
    const searchResults = context.document.body.search('ly', {matchSuffix: true});

    // Queue a command to load the search results and get the font property values.
    searchResults.load('font');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    await context.sync();
    console.log('Found count: ' + searchResults.items.length);

    // Queue a set of commands to change the font for each found item.
    for (let i = 0; i < searchResults.items.length; i++) {
        searchResults.items[i].font.color = 'orange';
        searchResults.items[i].font.highlightColor = 'black';
        searchResults.items[i].font.bold = true;
    }
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    await context.sync();
});  
```

```typescript
// Search using a wildcard
// Run a batch operation against the Word object model.
await Word.run(async (context) => {
    // Queue a command to search the document with a wildcard
    // for any string of characters that starts with 'to' and ends with 'n'.
    const searchResults = context.document.body.search('to*n', {matchWildcards: true});

    // Queue a command to load the search results and get the font property values.
    searchResults.load('font');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    await context.sync();
    console.log('Found count: ' + searchResults.items.length);

    // Queue a set of commands to change the font for each found item.
    for (let i = 0; i < searchResults.items.length; i++) {
        searchResults.items[i].font.color = 'purple';
        searchResults.items[i].font.highlightColor = 'pink';
        searchResults.items[i].font.bold = true;
    }
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    await context.sync();
}); 
```

---

### load(propertyNames)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.SearchOptions;
```

Parameters:
- `propertyNames`: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns: [Word.SearchOptions](https://learn.microsoft.com/en-us/javascript/api/word/word.searchoptions)

---

### load(propertyNamesAndPaths)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
    select?: string;
    expand?: string;
}): Word.SearchOptions;
```

Parameters:
- `propertyNamesAndPaths`:  
  `{ select?: string; expand?: string; }`  
  `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

Returns: [Word.SearchOptions](https://learn.microsoft.com/en-us/javascript/api/word/word.searchoptions)

---

### newObject(context)

Create a new instance of the `Word.SearchOptions` object.

```typescript
static newObject(context: OfficeExtension.ClientRequestContext): Word.SearchOptions;
```

Parameters:
- `context`: [OfficeExtension.ClientRequestContext](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext)

Returns: [Word.SearchOptions](https://learn.microsoft.com/en-us/javascript/api/word/word.searchoptions)

---

### set(properties, options)

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.SearchOptionsUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

Parameters:
- `properties`: [Word.Interfaces.SearchOptionsUpdateData](https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.searchoptionsupdatedata)  
  A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
- `options`: [OfficeExtension.UpdateOptions](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.updateoptions)  
  Provides an option to suppress errors if the properties object tries to set any read-only properties.

Returns: void

---

### set(properties)

Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.SearchOptions): void;
```

Parameters:
- `properties`: [Word.SearchOptions](https://learn.microsoft.com/en-us/javascript/api/word/word.searchoptions)

Returns: void

---

### toJSON()

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.SearchOptions` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.SearchOptionsData`) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.SearchOptionsData;
```

Returns: [Word.Interfaces.SearchOptionsData](https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.searchoptionsdata)