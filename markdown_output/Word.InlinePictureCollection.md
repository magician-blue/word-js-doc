# Word.InlinePictureCollection

**Package:** `word`

**API Set:** WordApi 1.1 None

**Extends:** `OfficeExtension.ClientObject`

## Description

Contains a collection of [Word.InlinePicture](/en-us/javascript/api/word/word.inlinepicture) objects.

## Class Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/15-images/insert-and-get-pictures.yaml

// Gets the first image in the document.
await Word.run(async (context) => {
  const firstPicture: Word.InlinePicture = context.document.body.inlinePictures.getFirst();
  firstPicture.load("width, height, imageFormat");

  await context.sync();
  console.log(`Image dimensions: ${firstPicture.width} x ${firstPicture.height}`, `Image format: ${firstPicture.imageFormat}`);
  // Get the image encoded as Base64.
  const base64 = firstPicture.getBase64ImageSrc();

  await context.sync();
  console.log(base64.value);
});
```

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from an InlinePictureCollection to verify the connection between the add-in and Word application before performing operations on inline pictures.

```typescript
await Word.run(async (context) => {
    const inlinePictures = context.document.body.inlinePictures;
    
    // Access the request context associated with the collection
    const requestContext = inlinePictures.context;
    
    // Verify the context is valid by using it to load properties
    inlinePictures.load("items");
    await requestContext.sync();
    
    console.log(`Connected to Word. Found ${inlinePictures.items.length} inline pictures.`);
    console.log(`Request context type: ${requestContext.constructor.name}`);
});
```

---

### items

**Type:** `Word.InlinePicture[]`

Gets the loaded child items in this collection.

#### Examples

**Example**: Get all inline pictures in the document and log their count and dimensions to the console.

```typescript
await Word.run(async (context) => {
    // Get all inline pictures in the document body
    const inlinePictures = context.document.body.inlinePictures;
    
    // Load the items array with width and height properties
    inlinePictures.load("items");
    
    await context.sync();
    
    // Access the loaded items array
    const pictureItems = inlinePictures.items;
    
    console.log(`Found ${pictureItems.length} inline pictures`);
    
    // Iterate through each picture in the items array
    pictureItems.forEach((picture, index) => {
        picture.load("width, height");
    });
    
    await context.sync();
    
    pictureItems.forEach((picture, index) => {
        console.log(`Picture ${index + 1}: ${picture.width}x${picture.height} pixels`);
    });
});
```

---

## Methods

### getFirst

**Kind:** `read`

Gets the first inline image in this collection. Throws an ItemNotFound error if this collection is empty.

#### Signature

**Returns:** `Word.InlinePicture`

#### Examples

**Example**: Get the first inline picture in the document and resize it to 150x150 pixels

```typescript
await Word.run(async (context) => {
    // Get all inline pictures in the document body
    const inlinePictures = context.document.body.inlinePictures;
    
    // Get the first inline picture
    const firstPicture = inlinePictures.getFirst();
    
    // Resize the picture
    firstPicture.width = 150;
    firstPicture.height = 150;
    
    await context.sync();
    
    console.log("First inline picture resized successfully");
});
```

---

### getFirstOrNullObject

**Kind:** `read`

Gets the first inline image in this collection. If this collection is empty, then this method will return an object with its isNullObject property set to true. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

#### Signature

**Returns:** `Word.InlinePicture`

#### Examples

**Example**: Check if a document contains any inline images and display an alert with the first image's width, or notify the user if no images exist.

```typescript
await Word.run(async (context) => {
    const inlinePictures = context.document.body.inlinePictures;
    const firstPicture = inlinePictures.getFirstOrNullObject();
    firstPicture.load("width, isNullObject");
    
    await context.sync();
    
    if (firstPicture.isNullObject) {
        console.log("No inline images found in the document.");
    } else {
        console.log(`First inline image width: ${firstPicture.width} pixels`);
    }
});
```

---

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.InlinePictureCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.InlinePictureCollection`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.InlinePictureCollection`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `OfficeExtension.LoadOption` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.InlinePictureCollection`

#### Examples

**Example**: Load and display the width and height properties of all inline pictures in the document

```typescript
await Word.run(async (context) => {
    // Get all inline pictures in the document
    const inlinePictures = context.document.body.inlinePictures;
    
    // Load width and height properties for all inline pictures
    inlinePictures.load("width, height");
    
    await context.sync();
    
    // Display the properties
    for (let i = 0; i < inlinePictures.items.length; i++) {
        console.log(`Picture ${i + 1}: Width = ${inlinePictures.items[i].width}, Height = ${inlinePictures.items[i].height}`);
    }
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.InlinePictureCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.InlinePictureCollectionData) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

#### Signature

**Returns:** `Word.Interfaces.InlinePictureCollectionData`

#### Examples

**Example**: Export inline picture collection data to JSON format for logging or external storage

```typescript
await Word.run(async (context) => {
    // Get all inline pictures in the document body
    const inlinePictures = context.document.body.inlinePictures;
    
    // Load properties we want to include in the JSON output
    inlinePictures.load("width,height,altTextTitle,altTextDescription");
    
    await context.sync();
    
    // Convert the collection to a plain JavaScript object
    const picturesData = inlinePictures.toJSON();
    
    // Log or use the JSON data
    console.log("Inline Pictures Data:", JSON.stringify(picturesData, null, 2));
    console.log(`Total pictures found: ${picturesData.items.length}`);
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.InlinePictureCollection`

#### Examples

**Example**: Track inline pictures in a document to maintain references across multiple sync calls while modifying their properties

```typescript
await Word.run(async (context) => {
    // Get all inline pictures in the document
    const inlinePictures = context.document.body.inlinePictures;
    inlinePictures.load("items");
    await context.sync();

    // Track the collection to use it across multiple sync calls
    inlinePictures.track();

    // First sync - modify width
    await context.sync();
    
    if (inlinePictures.items.length > 0) {
        inlinePictures.items[0].width = 200;
    }

    // Second sync - modify height
    await context.sync();
    
    if (inlinePictures.items.length > 0) {
        inlinePictures.items[0].height = 150;
    }

    await context.sync();

    // Untrack when done to free up memory
    inlinePictures.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.InlinePictureCollection`

#### Examples

**Example**: Load inline pictures from the document, process them, then untrack the collection to free memory after use.

```typescript
await Word.run(async (context) => {
    // Get all inline pictures in the document body
    const inlinePictures = context.document.body.inlinePictures;
    
    // Track the collection for processing
    inlinePictures.load("items");
    await context.sync();
    
    // Process the pictures (e.g., log count)
    console.log(`Found ${inlinePictures.items.length} inline pictures`);
    
    // Untrack the collection to release memory
    inlinePictures.untrack();
    await context.sync();
});
```

---

## Source

- https://docs.microsoft.com/en-us/javascript/api/word/word.inlinepicturecollection
