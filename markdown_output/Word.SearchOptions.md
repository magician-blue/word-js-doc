# Word.SearchOptions

**Package:** `word`

**API Set:** WordApi 1.1

**Extends:** `OfficeExtension.ClientObject`

## Description

Specifies the options to be included in a search operation. To learn more about how to use search options in the Word JavaScript APIs, read Use search options to find text in your Word add-in.

## Class Examples

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

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from SearchOptions to verify the connection between the add-in and Word before performing a search operation

```typescript
await Word.run(async (context) => {
    // Create search options
    const searchOptions = context.document.body.search("example", {
        matchCase: false
    }).getFirst().searchOptions;
    
    // Access the request context from the SearchOptions object
    const requestContext = searchOptions.context;
    
    // Verify the context is valid by using it to load properties
    requestContext.document.load("saved");
    await requestContext.sync();
    
    console.log("Request context is active and connected to Word");
});
```

---

### ignorePunct

**Type:** `boolean`

**Since:** WordApi 1.1

Specifies a value that indicates whether to ignore all punctuation characters between words. Corresponds to the Ignore punctuation check box in the Find and Replace dialog box.

#### Examples

**Example**: Search for the phrase "end user" in the document while ignoring punctuation, so it matches variations like "end-user" or "end.user"

```typescript
await Word.run(async (context) => {
    const searchOptions = {
        ignorePunct: true
    };
    
    const searchResults = context.document.body.search("end user", searchOptions);
    searchResults.load("text");
    
    await context.sync();
    
    console.log(`Found ${searchResults.items.length} matches`);
    searchResults.items.forEach(result => {
        result.font.highlightColor = "yellow";
    });
    
    await context.sync();
});
```

---

### ignoreSpace

**Type:** `boolean`

**Since:** WordApi 1.1

Specifies a value that indicates whether to ignore all whitespace between words. Corresponds to the Ignore whitespace characters check box in the Find and Replace dialog box.

#### Examples

**Example**: Find all occurrences of "New York" in the document, ignoring extra spaces between the words (so it matches "New York", "New  York", "New   York", etc.)

```typescript
await Word.run(async (context) => {
    const searchOptions = {
        ignoreSpace: true
    };
    
    const searchResults = context.document.body.search("New York", searchOptions);
    searchResults.load("length");
    
    await context.sync();
    
    console.log(`Found ${searchResults.items.length} matches (ignoring extra spaces)`);
});
```

---

### matchCase

**Type:** `boolean`

**Since:** WordApi 1.1

Specifies a value that indicates whether to perform a case sensitive search. Corresponds to the Match case check box in the Find and Replace dialog box.

#### Examples

**Example**: Search for the text "API" in the document with case-sensitive matching enabled, so it only finds exact case matches and ignores "api" or "Api".

```typescript
await Word.run(async (context) => {
    const searchOptions = {
        matchCase: true
    };
    
    const searchResults = context.document.body.search("API", searchOptions);
    searchResults.load("text");
    
    await context.sync();
    
    console.log(`Found ${searchResults.items.length} case-sensitive matches for "API"`);
});
```

---

### matchPrefix

**Type:** `boolean`

**Since:** WordApi 1.1

Specifies a value that indicates whether to match words that begin with the search string. Corresponds to the Match prefix check box in the Find and Replace dialog box.

#### Examples

**Example**: Find all words in the document that begin with the prefix "micro" (like "Microsoft", "microphone", "microscope")

```typescript
await Word.run(async (context) => {
    const searchOptions = {
        matchPrefix: true
    };
    
    const searchResults = context.document.body.search("micro", searchOptions);
    searchResults.load("text");
    
    await context.sync();
    
    console.log(`Found ${searchResults.items.length} words starting with "micro"`);
    searchResults.items.forEach(result => {
        console.log(result.text);
    });
});
```

---

### matchSuffix

**Type:** `boolean`

**Since:** WordApi 1.1

Specifies a value that indicates whether to match words that end with the search string. Corresponds to the Match suffix check box in the Find and Replace dialog box.

#### Examples

**Example**: Find all words in the document that end with "ing" (e.g., "running", "walking", "testing")

```typescript
await Word.run(async (context) => {
    const searchOptions = {
        matchSuffix: true
    };
    
    const searchResults = context.document.body.search("ing", searchOptions);
    searchResults.load("text");
    
    await context.sync();
    
    console.log(`Found ${searchResults.items.length} words ending with "ing"`);
    
    // Highlight the results
    for (let i = 0; i < searchResults.items.length; i++) {
        searchResults.items[i].font.highlightColor = "yellow";
    }
    
    await context.sync();
});
```

---

### matchWholeWord

**Type:** `boolean`

**Since:** WordApi 1.1

Specifies a value that indicates whether to find operation only entire words, not text that is part of a larger word. Corresponds to the Find whole words only check box in the Find and Replace dialog box.

#### Examples

**Example**: Search for the word "test" in the document, but only match it when it appears as a complete word (not as part of "testing" or "contest")

```typescript
await Word.run(async (context) => {
    const searchOptions = {
        matchWholeWord: true
    };
    
    const searchResults = context.document.body.search("test", searchOptions);
    searchResults.load("text");
    
    await context.sync();
    
    console.log(`Found ${searchResults.items.length} whole word matches for "test"`);
    
    // Highlight the results
    searchResults.items.forEach(result => {
        result.font.highlightColor = "yellow";
    });
    
    await context.sync();
});
```

---

### matchWildcards

**Type:** `boolean`

**Since:** WordApi 1.1

Specifies a value that indicates whether the search will be performed using special search operators. Corresponds to the Use wildcards check box in the Find and Replace dialog box.

#### Examples

**Example**: Search for all email addresses in the document using wildcard pattern matching

```typescript
await Word.run(async (context) => {
    const searchOptions = {
        matchWildcards: true
    };
    
    // Wildcard pattern for email addresses: <*@*.?>
    const searchResults = context.document.body.search("<*@*.?>", searchOptions);
    searchResults.load("text");
    
    await context.sync();
    
    console.log(`Found ${searchResults.items.length} email addresses`);
    searchResults.items.forEach((result) => {
        console.log(result.text);
    });
});
```

---

## Methods

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.SearchOptionsLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.SearchOptions`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.SearchOptions`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.SearchOptions`

#### Examples

**Example**: // Ignore punctuation search

```typescript
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

**Example**: // Search based on a prefix

```typescript
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

**Example**: // Search based on a suffix

```typescript
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

**Example**: // Search using a wildcard

```typescript
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

### newObject

**Kind:** `create`

Create a new instance of the `Word.SearchOptions` object.

#### Signature

**Parameters:**
- `context`: `OfficeExtension.ClientRequestContext` (required)

**Returns:** `Word.SearchOptions`

#### Examples

**Example**: Create a search options object to configure a case-sensitive search that matches whole words only

```typescript
await Word.run(async (context) => {
    // Create a new SearchOptions instance
    const searchOptions = context.document.body.context.application.createSearchOptions();
    
    // Configure the search options
    searchOptions.matchCase = true;
    searchOptions.matchWholeWord = true;
    searchOptions.ignoreSpace = false;
    searchOptions.ignorePunct = false;
    
    // Use the search options to find text
    const searchResults = context.document.body.search("example", searchOptions);
    searchResults.load("text");
    
    await context.sync();
    
    console.log(`Found ${searchResults.items.length} matches`);
});
```

---

### set

**Kind:** `write`

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `Interfaces.SearchOptionsUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.SearchOptions` (required)

  **Returns:** `void`

#### Examples

**Example**: Configure search options to find text that matches whole words only and is case-sensitive

```typescript
await Word.run(async (context) => {
    const searchOptions = context.document.body.search("Report", {
        matchCase: false,
        matchWholeWord: false
    }).getFirst().searchOptions;
    
    // Use set() to configure multiple search option properties at once
    searchOptions.set({
        matchCase: true,
        matchWholeWord: true,
        ignoreSpace: false,
        ignorePunct: false
    });
    
    await context.sync();
    console.log("Search options configured successfully");
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.SearchOptions` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.SearchOptionsData`) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.SearchOptionsData`

#### Examples

**Example**: Serialize search options to JSON format for logging or debugging purposes, showing how to capture and display the current search configuration.

```typescript
await Word.run(async (context) => {
    // Create search options with specific settings
    const searchOptions = {
        ignorePunct: true,
        ignoreSpace: true,
        matchCase: false,
        matchPrefix: false,
        matchSuffix: false,
        matchWholeWord: true,
        matchWildcards: false
    };
    
    // Search for text using the options
    const results = context.document.body.search("example", searchOptions);
    results.load("items");
    
    await context.sync();
    
    // Get the search options from the results and convert to JSON
    const options = results.items[0]?.searchOptions;
    if (options) {
        options.load("*");
        await context.sync();
        
        // Convert to JSON for logging/debugging
        const optionsJSON = options.toJSON();
        console.log("Search options as JSON:", JSON.stringify(optionsJSON, null, 2));
    }
});
```

---

## Source

- https://learn.microsoft.com/en-us/javascript/api/word/word.searchoptions
- https://learn.microsoft.com/en-us/javascript/api/word
- https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientobject
