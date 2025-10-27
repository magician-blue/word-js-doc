# Word.CheckboxContentControl

**Package:** `word`

**API Set:** WordApi 1.7 None

**Extends:** `OfficeExtension.ClientObject`

## Description

The data specific to content controls of type CheckBox.

## Class Examples

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

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a checkbox content control to verify the connection between the add-in and Word application before performing operations

```typescript
await Word.run(async (context) => {
    // Get the first checkbox content control in the document
    const checkboxControls = context.document.contentControls.getByTypes([Word.ContentControlType.checkBox]);
    const checkboxControl = checkboxControls.getFirst();
    
    // Load the checkbox content control
    checkboxControl.load("checkboxContentControl");
    await context.sync();
    
    // Access the checkbox-specific data
    const checkboxData = checkboxControl.checkboxContentControl;
    
    // Access the request context from the checkbox content control
    const requestContext = checkboxData.context;
    
    // Verify the context is valid by using it to perform an operation
    console.log("Request context is connected:", requestContext !== null);
    
    // Use the same context to load additional properties
    checkboxData.load("isChecked");
    await requestContext.sync();
    
    console.log("Checkbox state:", checkboxData.isChecked);
});
```

---

### isChecked

**Type:** `boolean`

**Since:** WordApi 1.7

Specifies the current state of the checkbox.

#### Examples

**Example**: Check if a checkbox content control is currently checked and toggle its state

```typescript
await Word.run(async (context) => {
    // Get the first checkbox content control in the document
    const checkboxControl = context.document.contentControls.getByTypes([Word.ContentControlType.checkBox]).getFirst();
    const checkboxData = checkboxControl.checkboxContentControl;
    
    // Load the current state
    checkboxData.load("isChecked");
    await context.sync();
    
    // Toggle the checkbox state
    checkboxData.isChecked = !checkboxData.isChecked;
    
    await context.sync();
    
    console.log(`Checkbox is now ${checkboxData.isChecked ? 'checked' : 'unchecked'}`);
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
  - `options`: `Word.Interfaces.CheckboxContentControlLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.CheckboxContentControl`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.CheckboxContentControl`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string }` (optional)
    `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.CheckboxContentControl`

#### Examples

**Example**: Load and read the checked state of the first checkbox content control in the document

```typescript
await Word.run(async (context) => {
    // Get the first checkbox content control in the document
    const checkboxContentControl = context.document.contentControls
        .getByTypes([Word.ContentControlType.checkBox])
        .getFirst();
    
    // Get the checkbox-specific data
    const checkboxData = checkboxContentControl.checkboxContentControl;
    
    // Load the checked property
    checkboxData.load("checked");
    
    // Sync to execute the load command
    await context.sync();
    
    // Read the loaded property
    console.log("Checkbox is checked:", checkboxData.checked);
});
```

---

### set

**Kind:** `write`

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `Interfaces.CheckboxContentControlUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.CheckboxContentControl` (required)

  **Returns:** `void`

#### Examples

**Example**: Configure a checkbox content control by setting multiple properties at once, including checked state and appearance symbols

```typescript
await Word.run(async (context) => {
    // Get the first checkbox content control in the document
    const checkboxContentControl = context.document.contentControls.getFirstOrNullObject();
    await context.sync();
    
    if (!checkboxContentControl.isNullObject && checkboxContentControl.type === Word.ContentControlType.checkBox) {
        // Set multiple checkbox properties at once
        checkboxContentControl.checkboxContentControl.set({
            isChecked: true,
            checkedState: "☑",
            uncheckedState: "☐"
        });
        
        await context.sync();
        console.log("Checkbox properties updated successfully");
    }
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.CheckboxContentControl` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.CheckboxContentControlData`) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.CheckboxContentControlData`

#### Examples

**Example**: Serialize a checkbox content control to JSON format for logging or data transfer purposes

```typescript
await Word.run(async (context) => {
    // Get the first checkbox content control in the document
    const checkboxContentControl = context.document.contentControls.getFirstOrNullObject();
    checkboxContentControl.load("type");
    
    await context.sync();
    
    if (!checkboxContentControl.isNullObject && checkboxContentControl.type === Word.ContentControlType.checkBox) {
        // Load checkbox-specific properties
        const checkboxData = checkboxContentControl.checkboxContentControl;
        checkboxData.load("state");
        
        await context.sync();
        
        // Convert to JSON for serialization
        const jsonData = checkboxData.toJSON();
        
        // Log or use the JSON representation
        console.log("Checkbox data:", JSON.stringify(jsonData, null, 2));
        // Output example: { "state": true }
    }
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.CheckboxContentControl`

#### Examples

**Example**: Track a checkbox content control across multiple sync calls to monitor and update its checked state without encountering InvalidObjectPath errors

```typescript
await Word.run(async (context) => {
    // Get the first checkbox content control in the document
    const checkboxControl = context.document.contentControls
        .getByTypes([Word.ContentControlType.checkBox])
        .getFirst();
    
    checkboxControl.load("checkboxContentControl");
    await context.sync();
    
    // Track the checkbox to use it across multiple sync calls
    const checkbox = checkboxControl.checkboxContentControl;
    checkbox.track();
    
    // First sync: Check the current state
    checkbox.load("isChecked");
    await context.sync();
    
    console.log("Current state:", checkbox.isChecked);
    
    // Second sync: Toggle the checkbox state
    checkbox.isChecked = !checkbox.isChecked;
    await context.sync();
    
    console.log("New state:", checkbox.isChecked);
    
    // Clean up tracking when done
    checkbox.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

#### Signature

**Returns:** `Word.CheckboxContentControl`

#### Examples

**Example**: Get a checkbox content control, check its state, and then untrack it to free memory after you're done using it.

```typescript
await Word.run(async (context) => {
    // Get the first checkbox content control in the document
    const checkboxControl = context.document.contentControls.getFirstOrNullObject();
    checkboxControl.load("type");
    
    await context.sync();
    
    if (checkboxControl.isNullObject) {
        console.log("No content control found");
        return;
    }
    
    // Get the checkbox-specific data
    const checkboxData = checkboxControl.checkboxContentControl;
    checkboxData.load("isChecked");
    
    await context.sync();
    
    // Use the checkbox data
    console.log("Checkbox is checked:", checkboxData.isChecked);
    
    // Untrack the object to free memory
    checkboxData.untrack();
    
    await context.sync();
});
```

---

## Source

- https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/10-content-controls/insert-and-change-checkbox-content-control.yaml
