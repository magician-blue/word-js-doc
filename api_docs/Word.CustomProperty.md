# Word.CustomProperty class

Package: [word](/en-us/javascript/api/word)

Represents a custom property.

Extends: [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks

[API set: WordApi 1.3]

#### Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/30-properties/read-write-custom-document-properties.yaml

await Word.run(async (context) => {
    const properties: Word.CustomPropertyCollection = context.document.properties.customProperties;
    properties.load("key,type,value");

    await context.sync();
    for (let i = 0; i < properties.items.length; i++)
        console.log("Property Name:" + properties.items[i].key + "; Type=" + properties.items[i].type + "; Property Value=" + properties.items[i].value);
});
```

## Properties

- [context](#context)  
  The request context associated with the object. This connects the add-in's process to the Office host application's process.
- [key](#key)  
  Gets the key of the custom property.
- [type](#type)  
  Gets the value type of the custom property. Possible values are: String, Number, Date, Boolean.
- [value](#value)  
  Specifies the value of the custom property. Note that even though Word on the web and the docx file format allow these properties to be arbitrarily long, the desktop version of Word will truncate string values to 255 16-bit chars (possibly creating invalid unicode by breaking up a surrogate pair).

## Methods

- [delete()](#delete)  
  Deletes the custom property.
- [load(options)](#loadoptions)  
  Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- [load(propertyNames)](#loadpropertynames)  
  Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- [load(propertyNamesAndPaths)](#loadpropertynamesandpaths)  
  Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- [set(properties, options)](#setproperties-options)  
  Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
- [set(properties)](#setproperties)  
  Sets multiple properties on the object at the same time, based on an existing loaded object.
- [toJSON()](#tojson)  
  Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`.
- [track()](#track)  
  Track the object for automatic adjustment based on surrounding changes in the document.
- [untrack()](#untrack)  
  Release the memory associated with this object, if it has previously been tracked.

## Property Details

### context

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value: [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### key

Gets the key of the custom property.

```typescript
readonly key: string;
```

Property Value: string

Remarks: [API set: WordApi 1.3]

### type

Gets the value type of the custom property. Possible values are: String, Number, Date, Boolean.

```typescript
readonly type: Word.DocumentPropertyType | "String" | "Number" | "Date" | "Boolean";
```

Property Value: [Word.DocumentPropertyType](/en-us/javascript/api/word/word.documentpropertytype) | "String" | "Number" | "Date" | "Boolean"

Remarks: [API set: WordApi 1.3]

### value

Specifies the value of the custom property. Note that even though Word on the web and the docx file format allow these properties to be arbitrarily long, the desktop version of Word will truncate string values to 255 16-bit chars (possibly creating invalid unicode by breaking up a surrogate pair).

```typescript
value: any;
```

Property Value: any

Remarks: [API set: WordApi 1.3]

## Method Details

### delete

Deletes the custom property.

```typescript
delete(): void;
```

Returns: void

Remarks: [API set: WordApi 1.3]

### load(options)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(options?: Word.Interfaces.CustomPropertyLoadOptions): Word.CustomProperty;
```

Parameters:
- options: [Word.Interfaces.CustomPropertyLoadOptions](/en-us/javascript/api/word/word.interfaces.custompropertyloadoptions)  
  Provides options for which properties of the object to load.

Returns: [Word.CustomProperty](/en-us/javascript/api/word/word.customproperty)

### load(propertyNames)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.CustomProperty;
```

Parameters:
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns: [Word.CustomProperty](/en-us/javascript/api/word/word.customproperty)

### load(propertyNamesAndPaths)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.CustomProperty;
```

Parameters:
- propertyNamesAndPaths:  
  `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

Returns: [Word.CustomProperty](/en-us/javascript/api/word/word.customproperty)

### set(properties, options)

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.CustomPropertyUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

Parameters:
- properties: [Word.Interfaces.CustomPropertyUpdateData](/en-us/javascript/api/word/word.interfaces.custompropertyupdatedata)  
  A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
- options: [OfficeExtension.UpdateOptions](/en-us/javascript/api/office/officeextension.updateoptions)  
  Provides an option to suppress errors if the properties object tries to set any read-only properties.

Returns: void

### set(properties)

Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.CustomProperty): void;
```

Parameters:
- properties: [Word.CustomProperty](/en-us/javascript/api/word/word.customproperty)

Returns: void

### toJSON

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.CustomProperty` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.CustomPropertyData`) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.CustomPropertyData;
```

Returns: [Word.Interfaces.CustomPropertyData](/en-us/javascript/api/word/word.interfaces.custompropertydata)

### track

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.CustomProperty;
```

Returns: [Word.CustomProperty](/en-us/javascript/api/word/word.customproperty)

### untrack

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

```typescript
untrack(): Word.CustomProperty;
```

Returns: [Word.CustomProperty](/en-us/javascript/api/word/word.customproperty)