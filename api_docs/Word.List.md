# Word.List class

Package: [word](/en-us/javascript/api/word)

Contains a collection of [Word.Paragraph](/en-us/javascript/api/word/word.paragraph) objects.

Extends
- [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks

[ API set: WordApi 1.3 ]

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

## Properties

- context  
  The request context associated with the object. This connects the add-in's process to the Office host application's process.

- id  
  Gets the list's id.

- levelExistences  
  Checks whether each of the 9 levels exists in the list. A true value indicates the level exists, which means there is at least one list item at that level.

- levelTypes  
  Gets all 9 level types in the list. Each type can be 'Bullet', 'Number', or 'Picture'.

- paragraphs  
  Gets paragraphs in the list.

## Methods

- getLevelFont(level)  
  Gets the font of the bullet, number, or picture at the specified level in the list.

- getLevelParagraphs(level)  
  Gets the paragraphs that occur at the specified level in the list.

- getLevelPicture(level)  
  Gets the Base64-encoded string representation of the picture at the specified level in the list.

- getLevelString(level)  
  Gets the bullet, number, or picture at the specified level as a string.

- insertParagraph(paragraphText, insertLocation)  
  Inserts a paragraph at the specified location.

- load(options)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

- load(propertyNames)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

- load(propertyNamesAndPaths)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

- resetLevelFont(level, resetFontName)  
  Resets the font of the bullet, number, or picture at the specified level in the list.

- setLevelAlignment(level, alignment)  
  Sets the alignment of the bullet, number, or picture at the specified level in the list.

- setLevelBullet(level, listBullet, charCode, fontName)  
  Sets the bullet format at the specified level in the list. If the bullet is 'Custom', the charCode is required.

- setLevelIndents(level, textIndent, bulletNumberPictureIndent)  
  Sets the two indents of the specified level in the list.

- setLevelNumbering(level, listNumbering, formatString)  
  Sets the numbering format at the specified level in the list.

- setLevelPicture(level, base64EncodedImage)  
  Sets the picture at the specified level in the list.

- setLevelStartingNumber(level, startingNumber)  
  Sets the starting number at the specified level in the list. Default value is 1.

- toJSON()  
  Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.List object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.ListData) that contains shallow copies of any loaded child properties from the original object.

- track()  
  Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

- untrack()  
  Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

## Property Details

### context

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value
- [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

---

### id

Gets the list's id.

```typescript
readonly id: number;
```

Property Value
- number

Remarks  
[ API set: WordApi 1.3 ]

---

### levelExistences

Checks whether each of the 9 levels exists in the list. A true value indicates the level exists, which means there is at least one list item at that level.

```typescript
readonly levelExistences: boolean[];
```

Property Value
- boolean[]

Remarks  
[ API set: WordApi 1.3 ]

#### Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/20-lists/organize-list.yaml

// Gets information about the first list in the document.
await Word.run(async (context) => {
  const lists: Word.ListCollection = context.document.body.lists;
  lists.load("items");

  await context.sync();

  if (lists.items.length === 0) {
    console.warn("There are no lists in this document.");
    return;
  }
  
  // Get the first list.
  const list: Word.List = lists.getFirst();
  list.load("levelTypes,levelExistences");

  await context.sync();

  const levelTypes  = list.levelTypes;
  console.log("Level types of the first list:");
  for (let i = 0; i < levelTypes.length; i++) {
    console.log(`- Level ${i + 1} (index ${i}): ${levelTypes[i]}`);
  }

  const levelExistences = list.levelExistences;
  console.log("Level existences of the first list:");
  for (let i = 0; i < levelExistences.length; i++) {
    console.log(`- Level ${i + 1} (index ${i}): ${levelExistences[i]}`);
  }
});
```

---

### levelTypes

Gets all 9 level types in the list. Each type can be 'Bullet', 'Number', or 'Picture'.

```typescript
readonly levelTypes: Word.ListLevelType[];
```

Property Value
- [Word.ListLevelType](/en-us/javascript/api/word/word.listleveltype)[]

Remarks  
[ API set: WordApi 1.3 ]

#### Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/20-lists/organize-list.yaml

// Gets information about the first list in the document.
await Word.run(async (context) => {
  const lists: Word.ListCollection = context.document.body.lists;
  lists.load("items");

  await context.sync();

  if (lists.items.length === 0) {
    console.warn("There are no lists in this document.");
    return;
  }
  
  // Get the first list.
  const list: Word.List = lists.getFirst();
  list.load("levelTypes,levelExistences");

  await context.sync();

  const levelTypes  = list.levelTypes;
  console.log("Level types of the first list:");
  for (let i = 0; i < levelTypes.length; i++) {
    console.log(`- Level ${i + 1} (index ${i}): ${levelTypes[i]}`);
  }

  const levelExistences = list.levelExistences;
  console.log("Level existences of the first list:");
  for (let i = 0; i < levelExistences.length; i++) {
    console.log(`- Level ${i + 1} (index ${i}): ${levelExistences[i]}`);
  }
});
```

---

### paragraphs

Gets paragraphs in the list.

```typescript
readonly paragraphs: Word.ParagraphCollection;
```

Property Value
- [Word.ParagraphCollection](/en-us/javascript/api/word/word.paragraphcollection)

Remarks  
[ API set: WordApi 1.3 ]

## Method Details

### getLevelFont(level)

Gets the font of the bullet, number, or picture at the specified level in the list.

```typescript
getLevelFont(level: number): Word.Font;
```

Parameters
- level: number  
  Required. The level in the list.

Returns
- [Word.Font](/en-us/javascript/api/word/word.font)

Remarks  
[ API set: WordApiDesktop 1.1 ]

---

### getLevelParagraphs(level)

Gets the paragraphs that occur at the specified level in the list.

```typescript
getLevelParagraphs(level: number): Word.ParagraphCollection;
```

Parameters
- level: number  
  Required. The level in the list.

Returns
- [Word.ParagraphCollection](/en-us/javascript/api/word/word.paragraphcollection)

Remarks  
[ API set: WordApi 1.3 ]

---

### getLevelPicture(level)

Gets the Base64-encoded string representation of the picture at the specified level in the list.

```typescript
getLevelPicture(level: number): OfficeExtension.ClientResult<string>;
```

Parameters
- level: number  
  Required. The level in the list.

Returns
- [OfficeExtension.ClientResult](/en-us/javascript/api/office/officeextension.clientresult)<string>

Remarks  
[ API set: WordApiDesktop 1.1 ]

---

### getLevelString(level)

Gets the bullet, number, or picture at the specified level as a string.

```typescript
getLevelString(level: number): OfficeExtension.ClientResult<string>;
```

Parameters
- level: number  
  Required. The level in the list.

Returns
- [OfficeExtension.ClientResult](/en-us/javascript/api/office/officeextension.clientresult)<string>

Remarks  
[ API set: WordApi 1.3 ]

---

### insertParagraph(paragraphText, insertLocation)

Inserts a paragraph at the specified location.

```typescript
insertParagraph(
  paragraphText: string,
  insertLocation:
    | Word.InsertLocation.start
    | Word.InsertLocation.end
    | Word.InsertLocation.before
    | Word.InsertLocation.after
    | "Start"
    | "End"
    | "Before"
    | "After"
): Word.Paragraph;
```

Parameters
- paragraphText: string  
  Required. The paragraph text to be inserted.
- insertLocation: start | end | before | after | "Start" | "End" | "Before" | "After"  
  Required. The value must be 'Start', 'End', 'Before', or 'After'.

Returns
- [Word.Paragraph](/en-us/javascript/api/word/word.paragraph)

Remarks  
[ API set: WordApi 1.3 ]

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

---

### load(options)

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(options?: Word.Interfaces.ListLoadOptions): Word.List;
```

Parameters
- options: [Word.Interfaces.ListLoadOptions](/en-us/javascript/api/word/word.interfaces.listloadoptions)  
  Provides options for which properties of the object to load.

Returns
- [Word.List](/en-us/javascript/api/word/word.list)

---

### load(propertyNames)

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.List;
```

Parameters
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns
- [Word.List](/en-us/javascript/api/word/word.list)

---

### load(propertyNamesAndPaths)

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
  select?: string;
  expand?: string;
}): Word.List;
```

Parameters
- propertyNamesAndPaths: { select?: string; expand?: string }  
  propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

Returns
- [Word.List](/en-us/javascript/api/word/word.list)

---

### resetLevelFont(level, resetFontName)

Resets the font of the bullet, number, or picture at the specified level in the list.

```typescript
resetLevelFont(level: number, resetFontName?: boolean): void;
```

Parameters
- level: number  
  Required. The level in the list.
- resetFontName: boolean  
  Optional. Indicates whether to reset the font name. Default is false that indicates the font name is kept unchanged.

Returns
- void

Remarks  
[ API set: WordApiDesktop 1.1 ]

---

### setLevelAlignment(level, alignment)

Sets the alignment of the bullet, number, or picture at the specified level in the list.

```typescript
setLevelAlignment(level: number, alignment: Word.Alignment): void;
```

Parameters
- level: number  
  Required. The level in the list.
- alignment: [Word.Alignment](/en-us/javascript/api/word/word.alignment)  
  Required. The level alignment that must be 'Left', 'Centered', or 'Right'.

Returns
- void

Remarks  
[ API set: WordApi 1.3 ]

---

### setLevelAlignment(level, alignment)

Sets the alignment of the bullet, number, or picture at the specified level in the list.

```typescript
setLevelAlignment(
  level: number,
  alignment: "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified"
): void;
```

Parameters
- level: number  
  Required. The level in the list.
- alignment: "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified"  
  Required. The level alignment that must be 'Left', 'Centered', or 'Right'.

Returns
- void

Remarks  
[ API set: WordApi 1.3 ]

---

### setLevelBullet(level, listBullet, charCode, fontName)

Sets the bullet format at the specified level in the list. If the bullet is 'Custom', the charCode is required.

```typescript
setLevelBullet(
  level: number,
  listBullet: Word.ListBullet,
  charCode?: number,
  fontName?: string
): void;
```

Parameters
- level: number  
  Required. The level in the list.
- listBullet: [Word.ListBullet](/en-us/javascript/api/word/word.listbullet)  
  Required. The bullet.
- charCode: number  
  Optional. The bullet character's code value. Used only if the bullet is 'Custom'.
- fontName: string  
  Optional. The bullet's font name. Used only if the bullet is 'Custom'.

Returns
- void

Remarks  
[ API set: WordApi 1.3 ]

#### Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/20-lists/organize-list.yaml

// Inserts a list starting with the first paragraph then set numbering and bullet types of the list items.
await Word.run(async (context) => {
  const paragraphs: Word.ParagraphCollection = context.document.body.paragraphs;
  paragraphs.load("$none");

  await context.sync();

  // Use the first paragraph to start a new list.
  const list: Word.List = paragraphs.items[0].startNewList();
  list.load("$none");

  await context.sync();

  // To add new items to the list, use Start or End on the insertLocation parameter.
  list.insertParagraph("New list item at the start of the list", "Start");
  const paragraph: Word.Paragraph = list.insertParagraph("New list item at the end of the list (set to list level 5)", "End");

  // Set numbering for list level 1.
  list.setLevelNumbering(0, Word.ListNumbering.arabic);

  // Set bullet type for list level 5.
  list.setLevelBullet(4, Word.ListBullet.arrow);

  // Set list level for the last item in this list.
  paragraph.listItem.level = 4;

  list.load("levelTypes");

  await context.sync();
});
```

---

### setLevelBullet(level, listBullet, charCode, fontName)

Sets the bullet format at the specified level in the list. If the bullet is 'Custom', the charCode is required.

```typescript
setLevelBullet(
  level: number,
  listBullet: "Custom" | "Solid" | "Hollow" | "Square" | "Diamonds" | "Arrow" | "Checkmark",
  charCode?: number,
  fontName?: string
): void;
```

Parameters
- level: number  
  Required. The level in the list.
- listBullet: "Custom" | "Solid" | "Hollow" | "Square" | "Diamonds" | "Arrow" | "Checkmark"  
  Required. The bullet.
- charCode: number  
  Optional. The bullet character's code value. Used only if the bullet is 'Custom'.
- fontName: string  
  Optional. The bullet's font name. Used only if the bullet is 'Custom'.

Returns
- void

Remarks  
[ API set: WordApi 1.3 ]

---

### setLevelIndents(level, textIndent, bulletNumberPictureIndent)

Sets the two indents of the specified level in the list.

```typescript
setLevelIndents(level: number, textIndent: number, bulletNumberPictureIndent: number): void;
```

Parameters
- level: number  
  Required. The level in the list.
- textIndent: number  
  Required. The text indent in points. It is the same as paragraph left indent.
- bulletNumberPictureIndent: number  
  Required. The relative indent, in points, of the bullet, number, or picture. It is the same as paragraph first line indent.

Returns
- void

Remarks  
[ API set: WordApi 1.3 ]

---

### setLevelNumbering(level, listNumbering, formatString)

Sets the numbering format at the specified level in the list.

```typescript
setLevelNumbering(
  level: number,
  listNumbering: Word.ListNumbering,
  formatString?: Array<string | number>
): void;
```

Parameters
- level: number  
  Required. The level in the list.
- listNumbering: [Word.ListNumbering](/en-us/javascript/api/word/word.listnumbering)  
  Required. The ordinal format.
- formatString: Array<string | number>  
  Optional. The numbering string format defined as an array of strings and/or integers. Each integer is a level of number type that is higher than or equal to this level. For example, an array of ["(", level - 1, ".", level, ")"] can define the format of "(2.c)", where 2 is the parent's item number and c is this level's item number.

Returns
- void

Remarks  
[ API set: WordApi 1.3 ]

#### Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/20-lists/organize-list.yaml

// Inserts a list starting with the first paragraph then set numbering and bullet types of the list items.
await Word.run(async (context) => {
  const paragraphs: Word.ParagraphCollection = context.document.body.paragraphs;
  paragraphs.load("$none");

  await context.sync();

  // Use the first paragraph to start a new list.
  const list: Word.List = paragraphs.items[0].startNewList();
  list.load("$none");

  await context.sync();

  // To add new items to the list, use Start or End on the insertLocation parameter.
  list.insertParagraph("New list item at the start of the list", "Start");
  const paragraph: Word.Paragraph = list.insertParagraph("New list item at the end of the list (set to list level 5)", "End");

  // Set numbering for list level 1.
  list.setLevelNumbering(0, Word.ListNumbering.arabic);

  // Set bullet type for list level 5.
  list.setLevelBullet(4, Word.ListBullet.arrow);

  // Set list level for the last item in this list.
  paragraph.listItem.level = 4;

  list.load("levelTypes");

  await context.sync();
});
```

---

### setLevelNumbering(level, listNumbering, formatString)

Sets the numbering format at the specified level in the list.

```typescript
setLevelNumbering(
  level: number,
  listNumbering: "None" | "Arabic" | "UpperRoman" | "LowerRoman" | "UpperLetter" | "LowerLetter",
  formatString?: Array<string | number>
): void;
```

Parameters
- level: number  
  Required. The level in the list.
- listNumbering: "None" | "Arabic" | "UpperRoman" | "LowerRoman" | "UpperLetter" | "LowerLetter"  
  Required. The ordinal format.
- formatString: Array<string | number>  
  Optional. The numbering string format defined as an array of strings and/or integers. Each integer is a level of number type that is higher than or equal to this level. For example, an array of ["(", level - 1, ".", level, ")"] can define the format of "(2.c)", where 2 is the parent's item number and c is this level's item number.

Returns
- void

Remarks  
[ API set: WordApi 1.3 ]

---

### setLevelPicture(level, base64EncodedImage)

Sets the picture at the specified level in the list.

```typescript
setLevelPicture(level: number, base64EncodedImage?: string): void;
```

Parameters
- level: number  
  Required. The level in the list.
- base64EncodedImage: string  
  Optional. The Base64-encoded image to be set. If not given, the default picture is set.

Returns
- void

Remarks  
[ API set: WordApiDesktop 1.1 ]

---

### setLevelStartingNumber(level, startingNumber)

Sets the starting number at the specified level in the list. Default value is 1.

```typescript
setLevelStartingNumber(level: number, startingNumber: number): void;
```

Parameters
- level: number  
  Required. The level in the list.
- startingNumber: number  
  Required. The number to start with.

Returns
- void

Remarks  
[ API set: WordApi 1.3 ]

---

### toJSON()

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.List object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.ListData) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.ListData;
```

Returns
- [Word.Interfaces.ListData](/en-us/javascript/api/word/word.interfaces.listdata)

---

### track()

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.List;
```

Returns
- [Word.List](/en-us/javascript/api/word/word.list)

---

### untrack()

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.List;
```

Returns
- [Word.List](/en-us/javascript/api/word/word.list)