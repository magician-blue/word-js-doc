# Word.CommentContentRange class

Package: [word](/en-us/javascript/api/word)

Extends: [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks
[ API set: WordApi 1.4 ]

#### Examples
```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-comments.yaml

// Gets the range of the first comment in the selected content.
await Word.run(async (context) => {
  const comment: Word.Comment = context.document.getSelection().getComments().getFirstOrNullObject();
  comment.load("contentRange");
  const range: Word.Range = comment.getRange();
  range.load("text");
  await context.sync();

  if (comment.isNullObject) {
    console.warn("No comments in the selection, so no range to get.");
    return;
  }

  console.log(`Comment location: ${range.text}`);
  const contentRange: Word.CommentContentRange = comment.contentRange;
  console.log("Comment content range:", contentRange);
});
```

## Properties
- bold  
  Specifies a value that indicates whether the comment text is bold.
- context  
  The request context associated with the object. This connects the add-in's process to the Office host application's process.
- hyperlink  
  Gets the first hyperlink in the range, or sets a hyperlink on the range. All hyperlinks in the range are deleted when you set a new hyperlink on the range.
- isEmpty  
  Checks whether the range length is zero.
- italic  
  Specifies a value that indicates whether the comment text is italicized.
- strikeThrough  
  Specifies a value that indicates whether the comment text has a strikethrough.
- text  
  Gets the text of the comment range.
- underline  
  Specifies a value that indicates the comment text's underline type. 'None' if the comment text isn't underlined.

## Methods
- insertText(text, insertLocation)  
  Inserts text into at the specified location. Note: For the modern comment, the content range tracked across context turns to empty if any revision to the comment is posted through the UI.
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
  Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.CommentContentRange object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.CommentContentRangeData) that contains shallow copies of any loaded child properties from the original object.
- track()  
  Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
- untrack()  
  Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

## Property Details

### bold
Specifies a value that indicates whether the comment text is bold.

```typescript
bold: boolean;
```

- Property Value: boolean

Remarks  
[ API set: WordApi 1.4 ]

### context
The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

- Property Value: [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### hyperlink
Gets the first hyperlink in the range, or sets a hyperlink on the range. All hyperlinks in the range are deleted when you set a new hyperlink on the range.

```typescript
hyperlink: string;
```

- Property Value: string

Remarks  
[ API set: WordApi 1.4 ]

### isEmpty
Checks whether the range length is zero.

```typescript
readonly isEmpty: boolean;
```

- Property Value: boolean

Remarks  
[ API set: WordApi 1.4 ]

### italic
Specifies a value that indicates whether the comment text is italicized.

```typescript
italic: boolean;
```

- Property Value: boolean

Remarks  
[ API set: WordApi 1.4 ]

### strikeThrough
Specifies a value that indicates whether the comment text has a strikethrough.

```typescript
strikeThrough: boolean;
```

- Property Value: boolean

Remarks  
[ API set: WordApi 1.4 ]

### text
Gets the text of the comment range.

```typescript
readonly text: string;
```

- Property Value: string

Remarks  
[ API set: WordApi 1.4 ]

### underline
Specifies a value that indicates the comment text's underline type. 'None' if the comment text isn't underlined.

```typescript
underline: Word.UnderlineType | "Mixed" | "None" | "Hidden" | "DotLine" | "Single" | "Word" | "Double" | "Thick" | "Dotted" | "DottedHeavy" | "DashLine" | "DashLineHeavy" | "DashLineLong" | "DashLineLongHeavy" | "DotDashLine" | "DotDashLineHeavy" | "TwoDotDashLine" | "TwoDotDashLineHeavy" | "Wave" | "WaveHeavy" | "WaveDouble";
```

- Property Value: [Word.UnderlineType](/en-us/javascript/api/word/word.underlinetype) | "Mixed" | "None" | "Hidden" | "DotLine" | "Single" | "Word" | "Double" | "Thick" | "Dotted" | "DottedHeavy" | "DashLine" | "DashLineHeavy" | "DashLineLong" | "DashLineLongHeavy" | "DotDashLine" | "DotDashLineHeavy" | "TwoDotDashLine" | "TwoDotDashLineHeavy" | "Wave" | "WaveHeavy" | "WaveDouble"

Remarks  
[ API set: WordApi 1.4 ]

## Method Details

### insertText(text, insertLocation)
Inserts text into at the specified location. Note: For the modern comment, the content range tracked across context turns to empty if any revision to the comment is posted through the UI.

```typescript
insertText(text: string, insertLocation: Word.InsertLocation | "Replace" | "Start" | "End" | "Before" | "After"): Word.CommentContentRange;
```

- Parameters:
  - text: string  
    Required. The text to be inserted in to the CommentContentRange.
  - insertLocation: [Word.InsertLocation](/en-us/javascript/api/word/word.insertlocation) | "Replace" | "Start" | "End" | "Before" | "After"  
    Required. The value must be 'Replace', 'Start', 'End', 'Before', or 'After'.
- Returns: [Word.CommentContentRange](/en-us/javascript/api/word/word.commentcontentrange)

Remarks  
[ API set: WordApi 1.4 ]

### load(options)
Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(options?: Word.Interfaces.CommentContentRangeLoadOptions): Word.CommentContentRange;
```

- Parameters:
  - options: [Word.Interfaces.CommentContentRangeLoadOptions](/en-us/javascript/api/word/word.interfaces.commentcontentrangeloadoptions)  
    Provides options for which properties of the object to load.
- Returns: [Word.CommentContentRange](/en-us/javascript/api/word/word.commentcontentrange)

### load(propertyNames)
Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.CommentContentRange;
```

- Parameters:
  - propertyNames: string | string[]  
    A comma-delimited string or an array of strings that specify the properties to load.
- Returns: [Word.CommentContentRange](/en-us/javascript/api/word/word.commentcontentrange)

### load(propertyNamesAndPaths)
Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.CommentContentRange;
```

- Parameters:
  - propertyNamesAndPaths:  
    `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
- Returns: [Word.CommentContentRange](/en-us/javascript/api/word/word.commentcontentrange)

### set(properties, options)
Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.CommentContentRangeUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

- Parameters:
  - properties: [Word.Interfaces.CommentContentRangeUpdateData](/en-us/javascript/api/word/word.interfaces.commentcontentrangeupdatedata)  
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - options: [OfficeExtension.UpdateOptions](/en-us/javascript/api/office/officeextension.updateoptions)  
    Provides an option to suppress errors if the properties object tries to set any read-only properties.
- Returns: void

### set(properties)
Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.CommentContentRange): void;
```

- Parameters:
  - properties: [Word.CommentContentRange](/en-us/javascript/api/word/word.commentcontentrange)
- Returns: void

### toJSON()
Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.CommentContentRange` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.CommentContentRangeData`) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.CommentContentRangeData;
```

- Returns: [Word.Interfaces.CommentContentRangeData](/en-us/javascript/api/word/word.interfaces.commentcontentrangedata)

### track()
Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.CommentContentRange;
```

- Returns: [Word.CommentContentRange](/en-us/javascript/api/word/word.commentcontentrange)

### untrack()
Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

```typescript
untrack(): Word.CommentContentRange;
```

- Returns: [Word.CommentContentRange](/en-us/javascript/api/word/word.commentcontentrange)