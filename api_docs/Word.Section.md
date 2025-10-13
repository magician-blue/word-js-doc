# Word.Section class

Package: [word](/en-us/javascript/api/word)

Represents a section in a Word document.

Extends: [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks
[ API set: WordApi 1.1 ]

### Examples
```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/insert-section-breaks.yaml

// Inserts a section break on the next page.
await Word.run(async (context) => {
  const body: Word.Body = context.document.body;
  body.insertBreak(Word.BreakType.sectionNext, Word.InsertLocation.end);

  await context.sync();

  console.log("Inserted section break on next page.");
});
```

## Properties
- body  
  Gets the body object of the section. This doesn't include the header/footer and other section metadata.
- borders  
  Returns a BorderUniversalCollection object that represents all the borders in the section.
- context  
  The request context associated with the object. This connects the add-in's process to the Office host application's process.
- pageSetup  
  Returns a PageSetup object that's associated with the section.
- protectedForForms  
  Specifies if the section is protected for forms.

## Methods
- getFooter(type)  
  Gets one of the section's footers.
- getFooter(type)  
  Gets one of the section's footers.
- getHeader(type)  
  Gets one of the section's headers.
- getHeader(type)  
  Gets one of the section's headers.
- getNext()  
  Gets the next section. Throws an ItemNotFound error if this section is the last one.
- getNextOrNullObject()  
  Gets the next section. If this section is the last one, then this method will return an object with its isNullObject property set to true. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).
- load(options)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNames)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNamesAndPaths)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- set(properties, options)  
  Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
- set(properties)  
  Sets multiple properties on the object at the same time, based on an existing loaded object.
- toJSON()  
  Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.Section object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.SectionData) that contains shallow copies of any loaded child properties from the original object.
- track()  
  Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
- untrack()  
  Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

## Property Details

### body
Gets the body object of the section. This doesn't include the header/footer and other section metadata.

```typescript
readonly body: Word.Body;
```

Property Value: [Word.Body](/en-us/javascript/api/word/word.body)

Remarks  
[ API set: WordApi 1.1 ]

---

### borders
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a BorderUniversalCollection object that represents all the borders in the section.

```typescript
readonly borders: Word.BorderUniversalCollection;
```

Property Value: [Word.BorderUniversalCollection](/en-us/javascript/api/word/word.borderuniversalcollection)

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### context
The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value: [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

---

### pageSetup
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a PageSetup object that's associated with the section.

```typescript
readonly pageSetup: Word.PageSetup;
```

Property Value: [Word.PageSetup](/en-us/javascript/api/word/word.pagesetup)

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### protectedForForms
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the section is protected for forms.

```typescript
protectedForForms: boolean;
```

Property Value: boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

## Method Details

### getFooter(type)
Gets one of the section's footers.

```typescript
getFooter(type: Word.HeaderFooterType): Word.Body;
```

Parameters
- type: [Word.HeaderFooterType](/en-us/javascript/api/word/word.headerfootertype)  
  Required. The type of footer to return. This value must be: 'Primary', 'FirstPage', or 'EvenPages'.

Returns: [Word.Body](/en-us/javascript/api/word/word.body)

Remarks  
[ API set: WordApi 1.1 ]

Examples
```typescript
// Run a batch operation against the Word object model.
await Word.run(async (context) => {
    
    // Create a proxy sectionsCollection object.
    const mySections = context.document.sections;
    
    // Queue a command to load the sections.
    mySections.load('body/style');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    await context.sync();
        
    // Create a proxy object the primary footer of the first section.
    // Note that the footer is a body object.
    const myFooter = mySections.items[0].getFooter(Word.HeaderFooterType.primary);
    
    // Queue a command to insert text at the end of the footer.
    myFooter.insertText("This is a footer.", Word.InsertLocation.end);
    
    // Queue a command to wrap the header in a content control.
    myFooter.insertContentControl();
                            
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    await context.sync();
    console.log("Added a footer to the first section.");   
});  
```

---

### getFooter(type)
Gets one of the section's footers.

```typescript
getFooter(type: "Primary" | "FirstPage" | "EvenPages"): Word.Body;
```

Parameters
- type: "Primary" | "FirstPage" | "EvenPages"  
  Required. The type of footer to return. This value must be: 'Primary', 'FirstPage', or 'EvenPages'.

Returns: [Word.Body](/en-us/javascript/api/word/word.body)

Remarks  
[ API set: WordApi 1.1 ]

Examples
```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/25-paragraph/insert-header-and-footer.yaml

await Word.run(async (context) => {
  context.document.sections
    .getFirst()
    .getFooter("Primary")
    .insertParagraph("This is a primary footer.", "End");

  await context.sync();
});
```

---

### getHeader(type)
Gets one of the section's headers.

```typescript
getHeader(type: Word.HeaderFooterType): Word.Body;
```

Parameters
- type: [Word.HeaderFooterType](/en-us/javascript/api/word/word.headerfootertype)  
  Required. The type of header to return. This value must be: 'Primary', 'FirstPage', or 'EvenPages'.

Returns: [Word.Body](/en-us/javascript/api/word/word.body)

Remarks  
[ API set: WordApi 1.1 ]

Examples
```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/25-paragraph/insert-header-and-footer.yaml

await Word.run(async (context) => {
  context.document.sections
    .getFirst()
    .getHeader(Word.HeaderFooterType.primary)
    .insertParagraph("This is a primary header.", "End");

  await context.sync();
});
```

---

### getHeader(type)
Gets one of the section's headers.

```typescript
getHeader(type: "Primary" | "FirstPage" | "EvenPages"): Word.Body;
```

Parameters
- type: "Primary" | "FirstPage" | "EvenPages"  
  Required. The type of header to return. This value must be: 'Primary', 'FirstPage', or 'EvenPages'.

Returns: [Word.Body](/en-us/javascript/api/word/word.body)

Remarks  
[ API set: WordApi 1.1 ]

Examples
```typescript
// Run a batch operation against the Word object model.
await Word.run(async (context) => {
    
    // Create a proxy sectionsCollection object.
    const mySections = context.document.sections;
    
    // Queue a command to load the sections.
    mySections.load('body/style');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    await context.sync();
    
    // Create a proxy object the primary header of the first section.
    // Note that the header is a body object.
    const myHeader = mySections.items[0].getHeader("Primary");
    
    // Queue a command to insert text at the end of the header.
    myHeader.insertText("This is a header.", Word.InsertLocation.end);
    
    // Queue a command to wrap the header in a content control.
    myHeader.insertContentControl();
                            
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    await context.sync();
    console.log("Added a header to the first section.");
});  
```

---

### getNext()
Gets the next section. Throws an ItemNotFound error if this section is the last one.

```typescript
getNext(): Word.Section;
```

Returns: [Word.Section](/en-us/javascript/api/word/word.section)

Remarks  
[ API set: WordApi 1.3 ]

---

### getNextOrNullObject()
Gets the next section. If this section is the last one, then this method will return an object with its isNullObject property set to true. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
getNextOrNullObject(): Word.Section;
```

Returns: [Word.Section](/en-us/javascript/api/word/word.section)

Remarks  
[ API set: WordApi 1.3 ]

---

### load(options)
Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(options?: Word.Interfaces.SectionLoadOptions): Word.Section;
```

Parameters
- options: [Word.Interfaces.SectionLoadOptions](/en-us/javascript/api/word/word.interfaces.sectionloadoptions)  
  Provides options for which properties of the object to load.

Returns: [Word.Section](/en-us/javascript/api/word/word.section)

---

### load(propertyNames)
Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.Section;
```

Parameters
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns: [Word.Section](/en-us/javascript/api/word/word.section)

---

### load(propertyNamesAndPaths)
Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.Section;
```

Parameters
- propertyNamesAndPaths:  
  {
  select?: string;
  expand?: string;
  }  
  propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

Returns: [Word.Section](/en-us/javascript/api/word/word.section)

---

### set(properties, options)
Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.SectionUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

Parameters
- properties: [Word.Interfaces.SectionUpdateData](/en-us/javascript/api/word/word.interfaces.sectionupdatedata)  
  A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
- options: [OfficeExtension.UpdateOptions](/en-us/javascript/api/office/officeextension.updateoptions)  
  Provides an option to suppress errors if the properties object tries to set any read-only properties.

Returns: void

---

### set(properties)
Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.Section): void;
```

Parameters
- properties: [Word.Section](/en-us/javascript/api/word/word.section)

Returns: void

---

### toJSON()
Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.Section object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.SectionData) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.SectionData;
```

Returns: [Word.Interfaces.SectionData](/en-us/javascript/api/word/word.interfaces.sectiondata)

---

### track()
Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.Section;
```

Returns: [Word.Section](/en-us/javascript/api/word/word.section)

---

### untrack()
Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.Section;
```

Returns: [Word.Section](/en-us/javascript/api/word/word.section)