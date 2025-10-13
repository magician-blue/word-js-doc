# Word.CheckboxContentControl class

Package: [word](/en-us/javascript/api/word)

The data specific to content controls of type CheckBox.

Extends: [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks
[API set: WordApi 1.7](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples
```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/10-content-controls/insert-and-change-checkbox-content-control.yaml

// Toggles the isChecked property on all checkbox content controls.
await Word.run(async (context) => {
  let contentControls = context.document.getContentControls({
    types: [Word.ContentControlType.checkBox]
  });
  contentControls.load("items");

  await context.sync();

  const length = contentControls.items.length;
  console.log(`Number of checkbox content controls: ${length}`);

  if (length <= 0) {
    return;
  }

  const checkboxContentControls = [];
  for (let i = 0; i < length; i++) {
    let contentControl = contentControls.items[i];
    contentControl.load("id,checkboxContentControl/isChecked");
    checkboxContentControls.push(contentControl);
  }

  await context.sync();

  console.log("isChecked state before:");
  const updatedCheckboxContentControls = [];
  for (let i = 0; i < checkboxContentControls.length; i++) {
    const currentCheckboxContentControl = checkboxContentControls[i];
    const isCheckedBefore = currentCheckboxContentControl.checkboxContentControl.isChecked;
    console.log(`id: ${currentCheckboxContentControl.id} ... isChecked: ${isCheckedBefore}`);

    currentCheckboxContentControl.checkboxContentControl.isChecked = !isCheckedBefore;
    currentCheckboxContentControl.load("id,checkboxContentControl/isChecked");
    updatedCheckboxContentControls.push(currentCheckboxContentControl);
  }

  await context.sync();

  console.log("isChecked state after:");
  for (let i = 0; i < updatedCheckboxContentControls.length; i++) {
    const currentCheckboxContentControl = updatedCheckboxContentControls[i];
    console.log(
      `id: ${currentCheckboxContentControl.id} ... isChecked: ${currentCheckboxContentControl.checkboxContentControl.isChecked}`
    );
  }
});
```

## Properties
- context: The request context associated with the object. This connects the add-in's process to the Office host application's process.
- isChecked: Specifies the current state of the checkbox.

## Methods
- load(options): Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- load(propertyNames): Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- load(propertyNamesAndPaths): Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- set(properties, options): Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
- set(properties): Sets multiple properties on the object at the same time, based on an existing loaded object.
- toJSON(): Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.CheckboxContentControl` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.CheckboxContentControlData`) that contains shallow copies of any loaded child properties from the original object.
- track(): Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
- untrack(): Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

## Property Details

### context
The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value: [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### isChecked
Specifies the current state of the checkbox.

```typescript
isChecked: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.7](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Method Details

### load(options)
Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(options?: Word.Interfaces.CheckboxContentControlLoadOptions): Word.CheckboxContentControl;
```

Parameters:
- options: [Word.Interfaces.CheckboxContentControlLoadOptions](/en-us/javascript/api/word/word.interfaces.checkboxcontentcontrolloadoptions)  
  Provides options for which properties of the object to load.

Returns: [Word.CheckboxContentControl](/en-us/javascript/api/word/word.checkboxcontentcontrol)

### load(propertyNames)
Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.CheckboxContentControl;
```

Parameters:
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns: [Word.CheckboxContentControl](/en-us/javascript/api/word/word.checkboxcontentcontrol)

### load(propertyNamesAndPaths)
Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
    select?: string;
    expand?: string;
}): Word.CheckboxContentControl;
```

Parameters:
- propertyNamesAndPaths:  
  select?: string; expand?: string  
  `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

Returns: [Word.CheckboxContentControl](/en-us/javascript/api/word/word.checkboxcontentcontrol)

### set(properties, options)
Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.CheckboxContentControlUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

Parameters:
- properties: [Word.Interfaces.CheckboxContentControlUpdateData](/en-us/javascript/api/word/word.interfaces.checkboxcontentcontrolupdatedata)  
  A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
- options: [OfficeExtension.UpdateOptions](/en-us/javascript/api/office/officeextension.updateoptions)  
  Provides an option to suppress errors if the properties object tries to set any read-only properties.

Returns: void

### set(properties)
Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.CheckboxContentControl): void;
```

Parameters:
- properties: [Word.CheckboxContentControl](/en-us/javascript/api/word/word.checkboxcontentcontrol)

Returns: void

### toJSON()
Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.CheckboxContentControl` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.CheckboxContentControlData`) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.CheckboxContentControlData;
```

Returns: [Word.Interfaces.CheckboxContentControlData](/en-us/javascript/api/word/word.interfaces.checkboxcontentcontroldata)

### track()
Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.CheckboxContentControl;
```

Returns: [Word.CheckboxContentControl](/en-us/javascript/api/word/word.checkboxcontentcontrol)

### untrack()
Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

```typescript
untrack(): Word.CheckboxContentControl;
```

Returns: [Word.CheckboxContentControl](/en-us/javascript/api/word/word.checkboxcontentcontrol)