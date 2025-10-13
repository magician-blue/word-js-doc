# Word.Bookmark class

- Package: [word](/en-us/javascript/api/word)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents a single bookmark in a document, selection, or range. The `Bookmark` object is a member of the `Bookmark` collection. The [Word.BookmarkCollection](/en-us/javascript/api/word/word.bookmarkcollection) includes all the bookmarks listed in the Bookmark dialog box (Insert menu).

Extends: [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

## Properties

- context: The request context associated with the object. This connects the add-in's process to the Office host application's process.
- end: Specifies the ending character position of the bookmark.
- isColumn: Returns `true` if the bookmark is a table column.
- isEmpty: Returns `true` if the bookmark is empty.
- name: Returns the name of the `Bookmark` object.
- range: Returns a `Range` object that represents the portion of the document that's contained in the `Bookmark` object.
- start: Specifies the starting character position of the bookmark.
- storyType: Returns the story type for the bookmark.

## Methods

- copyTo(name): Copies this bookmark to the new bookmark specified in the `name` argument and returns a `Bookmark` object.
- delete(): Deletes the bookmark.
- load(options): Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- load(propertyNames): Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- load(propertyNamesAndPaths): Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- select(): Selects the bookmark.
- set(properties, options): Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
- set(properties): Sets multiple properties on the object at the same time, based on an existing loaded object.
- toJSON(): Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.Bookmark` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.BookmarkData`) that contains shallow copies of any loaded child properties from the original object.
- track(): Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
- untrack(): Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

## Property Details

### context

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property value: [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

---

### end

Specifies the ending character position of the bookmark.

```typescript
end: number;
```

Property value: number

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

---

### isColumn

Returns `true` if the bookmark is a table column.

```typescript
readonly isColumn: boolean;
```

Property value: boolean

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

---

### isEmpty

Returns `true` if the bookmark is empty.

```typescript
readonly isEmpty: boolean;
```

Property value: boolean

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

---

### name

Returns the name of the `Bookmark` object.

```typescript
readonly name: string;
```

Property value: string

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

---

### range

Returns a `Range` object that represents the portion of the document that's contained in the `Bookmark` object.

```typescript
readonly range: Word.Range;
```

Property value: [Word.Range](/en-us/javascript/api/word/word.range)

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

---

### start

Specifies the starting character position of the bookmark.

```typescript
start: number;
```

Property value: number

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

---

### storyType

Returns the story type for the bookmark.

```typescript
readonly storyType: Word.StoryType | "MainText" | "Footnotes" | "Endnotes" | "Comments" | "TextFrame" | "EvenPagesHeader" | "PrimaryHeader" | "EvenPagesFooter" | "PrimaryFooter" | "FirstPageHeader" | "FirstPageFooter" | "FootnoteSeparator" | "FootnoteContinuationSeparator" | "FootnoteContinuationNotice" | "EndnoteSeparator" | "EndnoteContinuationSeparator" | "EndnoteContinuationNotice";
```

Property value: [Word.StoryType](/en-us/javascript/api/word/word.storytype) | "MainText" | "Footnotes" | "Endnotes" | "Comments" | "TextFrame" | "EvenPagesHeader" | "PrimaryHeader" | "EvenPagesFooter" | "PrimaryFooter" | "FirstPageHeader" | "FirstPageFooter" | "FootnoteSeparator" | "FootnoteContinuationSeparator" | "FootnoteContinuationNotice" | "EndnoteSeparator" | "EndnoteContinuationSeparator" | "EndnoteContinuationNotice"

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

## Method Details

### copyTo(name)

Copies this bookmark to the new bookmark specified in the `name` argument and returns a `Bookmark` object.

```typescript
copyTo(name: string): Word.Bookmark;
```

Parameters:
- name: string  
  The name of the new bookmark.

Returns: [Word.Bookmark](/en-us/javascript/api/word/word.bookmark)

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

---

### delete()

Deletes the bookmark.

```typescript
delete(): void;
```

Returns: void

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

---

### load(options)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(options?: Word.Interfaces.BookmarkLoadOptions): Word.Bookmark;
```

Parameters:
- options: [Word.Interfaces.BookmarkLoadOptions](/en-us/javascript/api/word/word.interfaces.bookmarkloadoptions)  
  Provides options for which properties of the object to load.

Returns: [Word.Bookmark](/en-us/javascript/api/word/word.bookmark)

---

### load(propertyNames)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.Bookmark;
```

Parameters:
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns: [Word.Bookmark](/en-us/javascript/api/word/word.bookmark)

---

### load(propertyNamesAndPaths)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.Bookmark;
```

Parameters:
- propertyNamesAndPaths:  
  `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

Returns: [Word.Bookmark](/en-us/javascript/api/word/word.bookmark)

---

### select()

Selects the bookmark.

```typescript
select(): void;
```

Returns: void

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

---

### set(properties, options)

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.BookmarkUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

Parameters:
- properties: [Word.Interfaces.BookmarkUpdateData](/en-us/javascript/api/word/word.interfaces.bookmarkupdatedata)  
  A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
- options: [OfficeExtension.UpdateOptions](/en-us/javascript/api/office/officeextension.updateoptions)  
  Provides an option to suppress errors if the properties object tries to set any read-only properties.

Returns: void

---

### set(properties)

Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.Bookmark): void;
```

Parameters:
- properties: [Word.Bookmark](/en-us/javascript/api/word/word.bookmark)

Returns: void

---

### toJSON()

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.Bookmark` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.BookmarkData`) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.BookmarkData;
```

Returns: [Word.Interfaces.BookmarkData](/en-us/javascript/api/word/word.interfaces.bookmarkdata)

---

### track()

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.Bookmark;
```

Returns: [Word.Bookmark](/en-us/javascript/api/word/word.bookmark)

---

### untrack()

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

```typescript
untrack(): Word.Bookmark;
```

Returns: [Word.Bookmark](/en-us/javascript/api/word/word.bookmark)