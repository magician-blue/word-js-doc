# Word.GlowFormat class

Package: [word](/en-us/javascript/api/word)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents the glow formatting for the font used by the range of text.

Extends
[OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks

[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

| Property | Description |
|---|---|
| [color](#word-word-glowformat-color-member) | Returns a `ColorFormat` object that represents the color for a glow effect. |
| [context](#word-word-glowformat-context-member) | The request context associated with the object. This connects the add-in's process to the Office host application's process. |
| [radius](#word-word-glowformat-radius-member) | Specifies the length of the radius for a glow effect. |
| [transparency](#word-word-glowformat-transparency-member) | Specifies the degree of transparency for the glow effect as a value between 0.0 (opaque) and 1.0 (clear). |

## Methods

| Method | Description |
|---|---|
| [load(options)](#word-word-glowformat-load-member(1)) | Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties. |
| [load(propertyNames)](#word-word-glowformat-load-member(2)) | Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties. |
| [load(propertyNamesAndPaths)](#word-word-glowformat-load-member(3)) | Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties. |
| [set(properties, options)](#word-word-glowformat-set-member(1)) | Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type. |
| [set(properties)](#word-word-glowformat-set-member(2)) | Sets multiple properties on the object at the same time, based on an existing loaded object. |
| [toJSON()](#word-word-glowformat-tojson-member(1)) | Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.GlowFormat` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.GlowFormatData`) that contains shallow copies of any loaded child properties from the original object. |
| [track()](#word-word-glowformat-track-member(1)) | Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection. |
| [untrack()](#word-word-glowformat-untrack-member(1)) | Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect. |

## Property Details

### color
<a id="word-word-glowformat-color-member"></a>

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `ColorFormat` object that represents the color for a glow effect.

```typescript
readonly color: Word.ColorFormat;
```

#### Property Value
[Word.ColorFormat](/en-us/javascript/api/word/word.colorformat)

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### context
<a id="word-word-glowformat-context-member"></a>

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

#### Property Value
[Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

---

### radius
<a id="word-word-glowformat-radius-member"></a>

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the length of the radius for a glow effect.

```typescript
radius: number;
```

#### Property Value
number

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### transparency
<a id="word-word-glowformat-transparency-member"></a>

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the degree of transparency for the glow effect as a value between 0.0 (opaque) and 1.0 (clear).

```typescript
transparency: number;
```

#### Property Value
number

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Method Details

### load(options)
<a id="word-word-glowformat-load-member(1)"></a>

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(options?: Word.Interfaces.GlowFormatLoadOptions): Word.GlowFormat;
```

#### Parameters
- options  
  [Word.Interfaces.GlowFormatLoadOptions](/en-us/javascript/api/word/word.interfaces.glowformatloadoptions)

  Provides options for which properties of the object to load.

#### Returns
[Word.GlowFormat](/en-us/javascript/api/word/word.glowformat)

---

### load(propertyNames)
<a id="word-word-glowformat-load-member(2)"></a>

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.GlowFormat;
```

#### Parameters
- propertyNames  
  string | string[]

  A comma-delimited string or an array of strings that specify the properties to load.

#### Returns
[Word.GlowFormat](/en-us/javascript/api/word/word.glowformat)

---

### load(propertyNamesAndPaths)
<a id="word-word-glowformat-load-member(3)"></a>

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.GlowFormat;
```

#### Parameters
- propertyNamesAndPaths  
  {
  select?: string;
  expand?: string;
  }

  `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

#### Returns
[Word.GlowFormat](/en-us/javascript/api/word/word.glowformat)

---

### set(properties, options)
<a id="word-word-glowformat-set-member(1)"></a>

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.GlowFormatUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

#### Parameters
- properties  
  [Word.Interfaces.GlowFormatUpdateData](/en-us/javascript/api/word/word.interfaces.glowformatupdatedata)

  A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.

- options  
  [OfficeExtension.UpdateOptions](/en-us/javascript/api/office/officeextension.updateoptions)

  Provides an option to suppress errors if the properties object tries to set any read-only properties.

#### Returns
void

---

### set(properties)
<a id="word-word-glowformat-set-member(2)"></a>

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.GlowFormat): void;
```

#### Parameters
- properties  
  [Word.GlowFormat](/en-us/javascript/api/word/word.glowformat)

#### Returns
void

---

### toJSON()
<a id="word-word-glowformat-tojson-member(1)"></a>

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.GlowFormat` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.GlowFormatData`) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.GlowFormatData;
```

#### Returns
[Word.Interfaces.GlowFormatData](/en-us/javascript/api/word/word.interfaces.glowformatdata)

---

### track()
<a id="word-word-glowformat-track-member(1)"></a>

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.GlowFormat;
```

#### Returns
[Word.GlowFormat](/en-us/javascript/api/word/word.glowformat)

---

### untrack()
<a id="word-word-glowformat-untrack-member(1)"></a>

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

```typescript
untrack(): Word.GlowFormat;
```

#### Returns
[Word.GlowFormat](/en-us/javascript/api/word/word.glowformat)