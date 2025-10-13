# Word.ParagraphFormat class

Package: [word](/en-us/javascript/api/word)

Represents a style of paragraph in a document.

Extends: [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks

[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-styles.yaml

// Sets certain aspects of the specified style's paragraph format e.g., the left indent size and the alignment.
await Word.run(async (context) => {
  const styleName = (document.getElementById("style-name") as HTMLInputElement).value;
  if (styleName == "") {
    console.warn("Enter a style name to update its paragraph format.");
    return;
  }

  const style: Word.Style = context.document.getStyles().getByNameOrNullObject(styleName);
  style.load();
  await context.sync();

  if (style.isNullObject) {
    console.warn(`There's no existing style with the name '${styleName}'.`);
  } else {
    style.paragraphFormat.leftIndent = 30;
    style.paragraphFormat.alignment = Word.Alignment.centered;
    console.log(`Successfully the paragraph format of the '${styleName}' style.`);
  }
});
```

## Properties

- alignment — Specifies the alignment for the specified paragraphs.
- context — The request context associated with the object. This connects the add-in's process to the Office host application's process.
- firstLineIndent — Specifies the value (in points) for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent.
- keepTogether — Specifies whether all lines in the specified paragraphs remain on the same page when Microsoft Word repaginates the document.
- keepWithNext — Specifies whether the specified paragraph remains on the same page as the paragraph that follows it when Microsoft Word repaginates the document.
- leftIndent — Specifies the left indent.
- lineSpacing — Specifies the line spacing (in points) for the specified paragraphs.
- lineUnitAfter — Specifies the amount of spacing (in gridlines) after the specified paragraphs.
- lineUnitBefore — Specifies the amount of spacing (in gridlines) before the specified paragraphs.
- mirrorIndents — Specifies whether left and right indents are the same width.
- outlineLevel — Specifies the outline level for the specified paragraphs.
- rightIndent — Specifies the right indent (in points) for the specified paragraphs.
- spaceAfter — Specifies the amount of spacing (in points) after the specified paragraph or text column.
- spaceBefore — Specifies the spacing (in points) before the specified paragraphs.
- widowControl — Specifies whether the first and last lines in the specified paragraph remain on the same page as the rest of the paragraph when Microsoft Word repaginates the document.

## Methods

- load(options) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNames) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNamesAndPaths) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- set(properties, options) — Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
- set(properties) — Sets multiple properties on the object at the same time, based on an existing loaded object.
- toJSON() — Overrides the JavaScript toJSON() method to provide more useful output when an API object is passed to JSON.stringify().
- track() — Track the object for automatic adjustment based on surrounding changes in the document.
- untrack() — Release the memory associated with this object, if it has previously been tracked.

## Property Details

### alignment

Specifies the alignment for the specified paragraphs.

```typescript
alignment: Word.Alignment | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified";
```

Property Value: [Word.Alignment](/en-us/javascript/api/word/word.alignment) | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified"

Remarks: [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-styles.yaml

// Sets certain aspects of the specified style's paragraph format e.g., the left indent size and the alignment.
await Word.run(async (context) => {
  const styleName = (document.getElementById("style-name") as HTMLInputElement).value;
  if (styleName == "") {
    console.warn("Enter a style name to update its paragraph format.");
    return;
  }

  const style: Word.Style = context.document.getStyles().getByNameOrNullObject(styleName);
  style.load();
  await context.sync();

  if (style.isNullObject) {
    console.warn(`There's no existing style with the name '${styleName}'.`);
  } else {
    style.paragraphFormat.leftIndent = 30;
    style.paragraphFormat.alignment = Word.Alignment.centered;
    console.log(`Successfully the paragraph format of the '${styleName}' style.`);
  }
});
```

### context

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value: [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### firstLineIndent

Specifies the value (in points) for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent.

```typescript
firstLineIndent: number;
```

Property Value: number

Remarks: [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### keepTogether

Specifies whether all lines in the specified paragraphs remain on the same page when Microsoft Word repaginates the document.

```typescript
keepTogether: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### keepWithNext

Specifies whether the specified paragraph remains on the same page as the paragraph that follows it when Microsoft Word repaginates the document.

```typescript
keepWithNext: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### leftIndent

Specifies the left indent.

```typescript
leftIndent: number;
```

Property Value: number

Remarks: [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-styles.yaml

// Sets certain aspects of the specified style's paragraph format e.g., the left indent size and the alignment.
await Word.run(async (context) => {
  const styleName = (document.getElementById("style-name") as HTMLInputElement).value;
  if (styleName == "") {
    console.warn("Enter a style name to update its paragraph format.");
    return;
  }

  const style: Word.Style = context.document.getStyles().getByNameOrNullObject(styleName);
  style.load();
  await context.sync();

  if (style.isNullObject) {
    console.warn(`There's no existing style with the name '${styleName}'.`);
  } else {
    style.paragraphFormat.leftIndent = 30;
    style.paragraphFormat.alignment = Word.Alignment.centered;
    console.log(`Successfully the paragraph format of the '${styleName}' style.`);
  }
});
```

### lineSpacing

Specifies the line spacing (in points) for the specified paragraphs.

```typescript
lineSpacing: number;
```

Property Value: number

Remarks: [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### lineUnitAfter

Specifies the amount of spacing (in gridlines) after the specified paragraphs.

```typescript
lineUnitAfter: number;
```

Property Value: number

Remarks: [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### lineUnitBefore

Specifies the amount of spacing (in gridlines) before the specified paragraphs.

```typescript
lineUnitBefore: number;
```

Property Value: number

Remarks: [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### mirrorIndents

Specifies whether left and right indents are the same width.

```typescript
mirrorIndents: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### outlineLevel

Specifies the outline level for the specified paragraphs.

```typescript
outlineLevel: Word.OutlineLevel | "OutlineLevel1" | "OutlineLevel2" | "OutlineLevel3" | "OutlineLevel4" | "OutlineLevel5" | "OutlineLevel6" | "OutlineLevel7" | "OutlineLevel8" | "OutlineLevel9" | "OutlineLevelBodyText";
```

Property Value: [Word.OutlineLevel](/en-us/javascript/api/word/word.outlinelevel) | "OutlineLevel1" | "OutlineLevel2" | "OutlineLevel3" | "OutlineLevel4" | "OutlineLevel5" | "OutlineLevel6" | "OutlineLevel7" | "OutlineLevel8" | "OutlineLevel9" | "OutlineLevelBodyText"

Remarks: [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### rightIndent

Specifies the right indent (in points) for the specified paragraphs.

```typescript
rightIndent: number;
```

Property Value: number

Remarks: [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### spaceAfter

Specifies the amount of spacing (in points) after the specified paragraph or text column.

```typescript
spaceAfter: number;
```

Property Value: number

Remarks: [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### spaceBefore

Specifies the spacing (in points) before the specified paragraphs.

```typescript
spaceBefore: number;
```

Property Value: number

Remarks: [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### widowControl

Specifies whether the first and last lines in the specified paragraph remain on the same page as the rest of the paragraph when Microsoft Word repaginates the document.

```typescript
widowControl: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Method Details

### load(options)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(options?: Word.Interfaces.ParagraphFormatLoadOptions): Word.ParagraphFormat;
```

Parameters:
- options: [Word.Interfaces.ParagraphFormatLoadOptions](/en-us/javascript/api/word/word.interfaces.paragraphformatloadoptions)  
  Provides options for which properties of the object to load.

Returns: [Word.ParagraphFormat](/en-us/javascript/api/word/word.paragraphformat)

### load(propertyNames)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.ParagraphFormat;
```

Parameters:
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns: [Word.ParagraphFormat](/en-us/javascript/api/word/word.paragraphformat)

### load(propertyNamesAndPaths)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.ParagraphFormat;
```

Parameters:
- propertyNamesAndPaths:  
  {
  select?: string;
  expand?: string;
  }  
  `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

Returns: [Word.ParagraphFormat](/en-us/javascript/api/word/word.paragraphformat)

### set(properties, options)

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.ParagraphFormatUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

Parameters:
- properties: [Word.Interfaces.ParagraphFormatUpdateData](/en-us/javascript/api/word/word.interfaces.paragraphformatupdatedata)  
  A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
- options: [OfficeExtension.UpdateOptions](/en-us/javascript/api/office/officeextension.updateoptions)  
  Provides an option to suppress errors if the properties object tries to set any read-only properties.

Returns: void

### set(properties)

Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.ParagraphFormat): void;
```

Parameters:
- properties: [Word.ParagraphFormat](/en-us/javascript/api/word/word.paragraphformat)

Returns: void

### toJSON()

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.ParagraphFormat` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.ParagraphFormatData`) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.ParagraphFormatData;
```

Returns: [Word.Interfaces.ParagraphFormatData](/en-us/javascript/api/word/word.interfaces.paragraphformatdata)

### track()

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.ParagraphFormat;
```

Returns: [Word.ParagraphFormat](/en-us/javascript/api/word/word.paragraphformat)

### untrack()

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

```typescript
untrack(): Word.ParagraphFormat;
```

Returns: [Word.ParagraphFormat](/en-us/javascript/api/word/word.paragraphformat)