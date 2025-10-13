# Word.XmlMapping class

Package: [word](/en-us/javascript/api/word)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents the XML mapping on a [Word.ContentControl](/en-us/javascript/api/word/word.contentcontrol) object between custom XML and that content control. An XML mapping is a link between the text in a content control and an XML element in the custom XML data store for this document.

Extends: [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

## Properties

- context  
  The request context associated with the object. This connects the add-in's process to the Office host application's process.
- customXmlNode  
  Returns a `CustomXmlNode` object that represents the custom XML node in the data store that the content control in the document maps to.
- customXmlPart  
  Returns a `CustomXmlPart` object that represents the custom XML part to which the content control in the document maps.
- isMapped  
  Returns whether the content control in the document is mapped to an XML node in the document's XML data store.
- prefixMappings  
  Returns the prefix mappings used to evaluate the XPath for the current XML mapping.
- xpath  
  Returns the XPath for the XML mapping, which evaluates to the currently mapped XML node.

## Methods

- delete()  
  Deletes the XML mapping from the parent content control.
- load(options)  
  Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- load(propertyNames)  
  Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- load(propertyNamesAndPaths)  
  Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- set(properties, options)  
  Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
- set(properties)  
  Sets multiple properties on the object at the same time, based on an existing loaded object.
- setMapping(xPath, options)  
  Allows creating or changing the XML mapping on the content control.
- setMappingByNode(node)  
  Allows creating or changing the XML data mapping on the content control.
- toJSON()  
  Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.XmlMapping` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.XmlMappingData`) that contains shallow copies of any loaded child properties from the original object.
- track()  
  Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
- untrack()  
  Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

## Property Details

### context

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value: [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### customXmlNode

Returns a `CustomXmlNode` object that represents the custom XML node in the data store that the content control in the document maps to.

```typescript
readonly customXmlNode: Word.CustomXmlNode;
```

Property Value: [Word.CustomXmlNode](/en-us/javascript/api/word/word.customxmlnode)

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

### customXmlPart

Returns a `CustomXmlPart` object that represents the custom XML part to which the content control in the document maps.

```typescript
readonly customXmlPart: Word.CustomXmlPart;
```

Property Value: [Word.CustomXmlPart](/en-us/javascript/api/word/word.customxmlpart)

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

### isMapped

Returns whether the content control in the document is mapped to an XML node in the document's XML data store.

```typescript
readonly isMapped: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

### prefixMappings

Returns the prefix mappings used to evaluate the XPath for the current XML mapping.

```typescript
readonly prefixMappings: string;
```

Property Value: string

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

### xpath

Returns the XPath for the XML mapping, which evaluates to the currently mapped XML node.

```typescript
readonly xpath: string;
```

Property Value: string

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

## Method Details

### delete()

Deletes the XML mapping from the parent content control.

```typescript
delete(): void;
```

Returns: void

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

### load(options)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(options?: Word.Interfaces.XmlMappingLoadOptions): Word.XmlMapping;
```

Parameters:
- options: [Word.Interfaces.XmlMappingLoadOptions](/en-us/javascript/api/word/word.interfaces.xmlmappingloadoptions)  
  Provides options for which properties of the object to load.

Returns: [Word.XmlMapping](/en-us/javascript/api/word/word.xmlmapping)

### load(propertyNames)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.XmlMapping;
```

Parameters:
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns: [Word.XmlMapping](/en-us/javascript/api/word/word.xmlmapping)

### load(propertyNamesAndPaths)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
    select?: string;
    expand?: string;
}): Word.XmlMapping;
```

Parameters:
- propertyNamesAndPaths:  
  `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

Returns: [Word.XmlMapping](/en-us/javascript/api/word/word.xmlmapping)

### set(properties, options)

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.XmlMappingUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

Parameters:
- properties: [Word.Interfaces.XmlMappingUpdateData](/en-us/javascript/api/word/word.interfaces.xmlmappingupdatedata)  
  A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
- options: [OfficeExtension.UpdateOptions](/en-us/javascript/api/office/officeextension.updateoptions)  
  Provides an option to suppress errors if the properties object tries to set any read-only properties.

Returns: void

### set(properties)

Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.XmlMapping): void;
```

Parameters:
- properties: [Word.XmlMapping](/en-us/javascript/api/word/word.xmlmapping)

Returns: void

### setMapping(xPath, options)

Allows creating or changing the XML mapping on the content control.

```typescript
setMapping(xPath: string, options?: Word.XmlSetMappingOptions): OfficeExtension.ClientResult<boolean>;
```

Parameters:
- xPath: string  
  The XPath expression to evaluate.
- options: [Word.XmlSetMappingOptions](/en-us/javascript/api/word/word.xmlsetmappingoptions)  
  Optional. The options available for setting the XML mapping.

Returns: [`OfficeExtension.ClientResult`](/en-us/javascript/api/office/officeextension.clientresult)<boolean>

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

### setMappingByNode(node)

Allows creating or changing the XML data mapping on the content control.

```typescript
setMappingByNode(node: Word.CustomXmlNode): OfficeExtension.ClientResult<boolean>;
```

Parameters:
- node: [Word.CustomXmlNode](/en-us/javascript/api/word/word.customxmlnode)  
  The custom XML node to map.

Returns: [`OfficeExtension.ClientResult`](/en-us/javascript/api/office/officeextension.clientresult)<boolean>

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

### toJSON()

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.XmlMapping` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.XmlMappingData`) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.XmlMappingData;
```

Returns: [Word.Interfaces.XmlMappingData](/en-us/javascript/api/word/word.interfaces.xmlmappingdata)

### track()

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.XmlMapping;
```

Returns: [Word.XmlMapping](/en-us/javascript/api/word/word.xmlmapping)

### untrack()

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

```typescript
untrack(): Word.XmlMapping;
```

Returns: [Word.XmlMapping](/en-us/javascript/api/word/word.xmlmapping)