# Word.CustomXmlPartCollection class

Package: [word](/en-us/javascript/api/word)

Contains the collection of [Word.CustomXmlPart](/en-us/javascript/api/word/word.customxmlpart) objects.

- Extends: [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks

[API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- [context](#word-word-customxmlpartcollection-context-member)
  - The request context associated with the object. This connects the add-in's process to the Office host application's process.
- [items](#word-word-customxmlpartcollection-items-member)
  - Gets the loaded child items in this collection.

## Methods

- [add(xml)](<#word-word-customxmlpartcollection-add-member(1)>)
  - Adds a new custom XML part to the document.
- [getByNamespace(namespaceUri)](<#word-word-customxmlpartcollection-getbynamespace-member(1)>)
  - Gets a new scoped collection of custom XML parts whose namespaces match the given namespace.
- [getCount()](<#word-word-customxmlpartcollection-getcount-member(1)>)
  - Gets the number of items in the collection.
- [getItem(id)](<#word-word-customxmlpartcollection-getitem-member(1)>)
  - Gets a custom XML part based on its ID.
- [getItemOrNullObject(id)](<#word-word-customxmlpartcollection-getitemornullobject-member(1)>)
  - Gets a custom XML part based on its ID. If the CustomXmlPart doesn't exist, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).
- [load(options)](<#word-word-customxmlpartcollection-load-member(1)>)
  - Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- [load(propertyNames)](<#word-word-customxmlpartcollection-load-member(2)>)
  - Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- [load(propertyNamesAndPaths)](<#word-word-customxmlpartcollection-load-member(3)>)
  - Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- [toJSON()](<#word-word-customxmlpartcollection-tojson-member(1)>)
  - Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.CustomXmlPartCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.CustomXmlPartCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
- [track()](<#word-word-customxmlpartcollection-track-member(1)>)
  - Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
- [untrack()](<#word-word-customxmlpartcollection-untrack-member(1)>)
  - Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

## Property Details

<a id="word-word-customxmlpartcollection-context-member"></a>
### context

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```
context: RequestContext;
```

- Property Value: [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

<a id="word-word-customxmlpartcollection-items-member"></a>
### items

Gets the loaded child items in this collection.

```
readonly items: Word.CustomXmlPart[];
```

- Property Value: [Word.CustomXmlPart](/en-us/javascript/api/word/word.customxmlpart)[]

## Method Details

<a id="word-word-customxmlpartcollection-add-member(1)"></a>
### add(xml)

Adds a new custom XML part to the document.

```
add(xml: string): Word.CustomXmlPart;
```

- Parameters:
  - xml (string): Required. XML content. Must be a valid XML fragment.
- Returns: [Word.CustomXmlPart](/en-us/javascript/api/word/word.customxmlpart)

Remarks:
[API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples

```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-custom-xml-part-ns.yaml

// Adds a custom XML part.
// If you want to populate the CustomXml.namespaceUri property, you must include the xmlns attribute.
await Word.run(async (context) => {
  const originalXml =
    "<Reviewers xmlns='http://schemas.contoso.com/review/1.0'><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>";
  const customXmlPart = context.document.customXmlParts.add(originalXml);
  customXmlPart.load(["id", "namespaceUri"]);
  const xmlBlob = customXmlPart.getXml();

  await context.sync();

  const readableXml = addLineBreaksToXML(xmlBlob.value);
  console.log(`Added custom XML part with namespace URI ${customXmlPart.namespaceUri}:`, readableXml);

  // Store the XML part's ID in a setting so the ID is available to other functions.
  const settings: Word.SettingCollection = context.document.settings;
  settings.add("ContosoReviewXmlPartIdNS", customXmlPart.id);

  await context.sync();
});

...

// Adds a custom XML part.
await Word.run(async (context) => {
  const originalXml =
    "<Reviewers><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>";
  const customXmlPart: Word.CustomXmlPart = context.document.customXmlParts.add(originalXml);
  customXmlPart.load("id");
  const xmlBlob = customXmlPart.getXml();

  await context.sync();

  const readableXml = addLineBreaksToXML(xmlBlob.value);
  console.log("Added custom XML part:", readableXml);

  // Store the XML part's ID in a setting so the ID is available to other functions.
  const settings: Word.SettingCollection = context.document.settings;
  settings.add("ContosoReviewXmlPartId", customXmlPart.id);

  await context.sync();
});
```

<a id="word-word-customxmlpartcollection-getbynamespace-member(1)"></a>
### getByNamespace(namespaceUri)

Gets a new scoped collection of custom XML parts whose namespaces match the given namespace.

```
getByNamespace(namespaceUri: string): Word.CustomXmlPartScopedCollection;
```

- Parameters:
  - namespaceUri (string): Required. The namespace URI.
- Returns: [Word.CustomXmlPartScopedCollection](/en-us/javascript/api/word/word.customxmlpartscopedcollection)

Remarks:
[API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples

```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-custom-xml-part-ns.yaml

// Original XML: <Reviewers xmlns='http://schemas.contoso.com/review/1.0'><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>

// Gets the custom XML parts with the specified namespace URI.
await Word.run(async (context) => {
  const namespaceUri = "http://schemas.contoso.com/review/1.0";
  console.log(`Specified namespace URI: ${namespaceUri}`);
  const scopedCustomXmlParts: Word.CustomXmlPartScopedCollection =
    context.document.customXmlParts.getByNamespace(namespaceUri);
  scopedCustomXmlParts.load("items");
  await context.sync();

  console.log(`Number of custom XML parts found with this namespace: ${!scopedCustomXmlParts.items ? 0 : scopedCustomXmlParts.items.length}`);
});
```

<a id="word-word-customxmlpartcollection-getcount-member(1)"></a>
### getCount()

Gets the number of items in the collection.

```
getCount(): OfficeExtension.ClientResult<number>;
```

- Returns: [OfficeExtension.ClientResult](/en-us/javascript/api/office/officeextension.clientresult)<number>

Remarks:
[API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-customxmlpartcollection-getitem-member(1)"></a>
### getItem(id)

Gets a custom XML part based on its ID.

```
getItem(id: string): Word.CustomXmlPart;
```

- Parameters:
  - id (string): ID or index of the custom XML part to be retrieved.
- Returns: [Word.CustomXmlPart](/en-us/javascript/api/word/word.customxmlpart)

Remarks:
[API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples

```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-custom-xml-part-ns.yaml

// Original XML: <Reviewers xmlns='http://schemas.contoso.com/review/1.0'><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>

// Queries a custom XML part for elements matching the search terms.
await Word.run(async (context) => {
  const settings: Word.SettingCollection = context.document.settings;
  const xmlPartIDSetting: Word.Setting = settings.getItemOrNullObject("ContosoReviewXmlPartIdNS").load("value");

  await context.sync();

  if (xmlPartIDSetting.value) {
    const customXmlPart: Word.CustomXmlPart = context.document.customXmlParts.getItem(xmlPartIDSetting.value);
    const xpathToQueryFor = "/contoso:Reviewers";
    const clientResult = customXmlPart.query(xpathToQueryFor, {
      contoso: "http://schemas.contoso.com/review/1.0"
    });

    await context.sync();

    console.log(`Queried custom XML part for ${xpathToQueryFor} and found ${clientResult.value.length} matches:`);
    for (let i = 0; i < clientResult.value.length; i++) {
      console.log(clientResult.value[i]);
    }
  } else {
    console.warn("Didn't find custom XML part to query.");
  }
});

...

// Original XML: <Reviewers><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>

// Queries a custom XML part for elements matching the search terms.
await Word.run(async (context) => {
  const settings: Word.SettingCollection = context.document.settings;
  const xmlPartIDSetting: Word.Setting = settings.getItemOrNullObject("ContosoReviewXmlPartId").load("value");

  await context.sync();

  if (xmlPartIDSetting.value) {
    const customXmlPart: Word.CustomXmlPart = context.document.customXmlParts.getItem(xmlPartIDSetting.value);
    const xpathToQueryFor = "/Reviewers/Reviewer";
    const clientResult = customXmlPart.query(xpathToQueryFor, {
      contoso: "http://schemas.contoso.com/review/1.0"
    });

    await context.sync();

    console.log(`Queried custom XML part for ${xpathToQueryFor} and found ${clientResult.value.length} matches:`);
    for (let i = 0; i < clientResult.value.length; i++) {
      console.log(clientResult.value[i]);
    }
  } else {
    console.warn("Didn't find custom XML part to query.");
  }
});
```

<a id="word-word-customxmlpartcollection-getitemornullobject-member(1)"></a>
### getItemOrNullObject(id)

Gets a custom XML part based on its ID. If the CustomXmlPart doesn't exist, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```
getItemOrNullObject(id: string): Word.CustomXmlPart;
```

- Parameters:
  - id (string): Required. ID of the object to be retrieved.
- Returns: [Word.CustomXmlPart](/en-us/javascript/api/word/word.customxmlpart)

Remarks:
[API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-customxmlpartcollection-load-member(1)"></a>
### load(options)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```
load(options?: Word.Interfaces.CustomXmlPartCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions): Word.CustomXmlPartCollection;
```

- Parameters:
  - options ([Word.Interfaces.CustomXmlPartCollectionLoadOptions](/en-us/javascript/api/word/word.interfaces.customxmlpartcollectionloadoptions) & [Word.Interfaces.CollectionLoadOptions](/en-us/javascript/api/word/word.interfaces.collectionloadoptions)): Provides options for which properties of the object to load.
- Returns: [Word.CustomXmlPartCollection](/en-us/javascript/api/word/word.customxmlpartcollection)

<a id="word-word-customxmlpartcollection-load-member(2)"></a>
### load(propertyNames)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```
load(propertyNames?: string | string[]): Word.CustomXmlPartCollection;
```

- Parameters:
  - propertyNames (string | string[]): A comma-delimited string or an array of strings that specify the properties to load.
- Returns: [Word.CustomXmlPartCollection](/en-us/javascript/api/word/word.customxmlpartcollection)

<a id="word-word-customxmlpartcollection-load-member(3)"></a>
### load(propertyNamesAndPaths)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```
load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Word.CustomXmlPartCollection;
```

- Parameters:
  - propertyNamesAndPaths ([OfficeExtension.LoadOption](/en-us/javascript/api/office/officeextension.loadoption)): `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
- Returns: [Word.CustomXmlPartCollection](/en-us/javascript/api/word/word.customxmlpartcollection)

<a id="word-word-customxmlpartcollection-tojson-member(1)"></a>
### toJSON()

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.CustomXmlPartCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.CustomXmlPartCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

```
toJSON(): Word.Interfaces.CustomXmlPartCollectionData;
```

- Returns: [Word.Interfaces.CustomXmlPartCollectionData](/en-us/javascript/api/word/word.interfaces.customxmlpartcollectiondata)

<a id="word-word-customxmlpartcollection-track-member(1)"></a>
### track()

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```
track(): Word.CustomXmlPartCollection;
```

- Returns: [Word.CustomXmlPartCollection](/en-us/javascript/api/word/word.customxmlpartcollection)

<a id="word-word-customxmlpartcollection-untrack-member(1)"></a>
### untrack()

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

```
untrack(): Word.CustomXmlPartCollection;
```

- Returns: [Word.CustomXmlPartCollection](/en-us/javascript/api/word/word.customxmlpartcollection)