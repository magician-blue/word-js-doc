# Word.ListTemplate class

Package: [word](/en-us/javascript/api/word)

Represents a list template.

Extends
- [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks

[API set: WordApiDesktop 1.1]

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

- [context](#word-word-listtemplate-context-member)  
  The request context associated with the object. This connects the add-in's process to the Office host application's process.

- [listLevels](#word-word-listtemplate-listlevels-member)  
  Gets a ListLevelCollection object that represents all the levels for the list template.

- [outlineNumbered](#word-word-listtemplate-outlinenumbered-member)  
  Specifies whether the list template is outline numbered.

## Methods

- [load(options)](#word-word-listtemplate-load-member1)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

- [load(propertyNames)](#word-word-listtemplate-load-member2)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

- [load(propertyNamesAndPaths)](#word-word-listtemplate-load-member3)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

- [set(properties, options)](#word-word-listtemplate-set-member1)  
  Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

- [set(properties)](#word-word-listtemplate-set-member2)  
  Sets multiple properties on the object at the same time, based on an existing loaded object.

- [toJSON()](#word-word-listtemplate-tojson-member1)  
  Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.ListTemplate object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.ListTemplateData) that contains shallow copies of any loaded child properties from the original object.

- [track()](#word-word-listtemplate-track-member1)  
  Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to