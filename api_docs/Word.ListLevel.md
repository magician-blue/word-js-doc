# Word.ListLevel class

Package: https://learn.microsoft.com/en-us/javascript/api/word

Represents a list level.

Extends
- https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientobject

## Remarks

[API set: WordApiDesktop 1.1](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/20-lists/manage-list-styles.yaml

// Gets the properties of the specified style.
await Word.run(async (context) => {
  const styleName = (document.getElementById("style-name-to-use") as HTMLInputElement).value;
  if (styleName == "") {
    console.warn("Enter a style name to get properties.");
    return;
  }

  const style: Word.Style = context.document.getStyles().getByNameOrNullObject(styleName);
  style.load("type");
  await context.sync();

  if (style.isNullObject || style.type != Word.StyleType.list) {
    console.warn(`There's no existing style with the name '${styleName}'. Or this isn't a list style.`);
  } else {
    // Load objects to log properties and their values in the console.
    style.load();
    style.listTemplate.load();
    await context.sync();

    console.log(`Properties of the '${styleName}' style:`, style);

    const listLevels = style.listTemplate.listLevels;
    listLevels.load("items");
    await context.sync();

    console.log(`List levels of the '${styleName}' style:`, listLevels);
  }
});
```

## Properties

- alignment  
  Specifies the horizontal alignment of the list level. The value can be 'Left', 'Centered', or 'Right'.
- context  
  The request context associated with the object. This connects the add-in's process to the Office host application's process.
- font  
  Gets a Font object that represents the character formatting of the specified object.
- linkedStyle  
  Specifies the name of the style that's linked to the specified list level object.
- numberFormat  
  Specifies the number format for the specified list level.
- numberPosition  
  Specifies the position (in points) of the number or bullet for the specified list level object.
- numberStyle  
  Specifies the number style for the list level object.
- resetOnHigher  
  Specifies the list level that must appear before the specified list level restarts numbering at 1.
- startAt  
  Specifies the starting number for the specified list level object.
- tabPosition  
  Specifies the tab position for the specified list level object.
- textPosition  
  Specifies the position (in points) for the second line of wrapping text for the specified list level object.
- trailingCharacter  
  Specifies the character inserted after the number for the specified list level.

## Methods

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
  Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.ListLevel object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.ListLevelData) that contains shallow copies of any loaded child properties from the original object.
- track()  
  Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the par