# Word.DatePickerContentControl class

Package: word

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents the DatePickerContentControl object.

Extends: OfficeExtension.ClientObject

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

## Properties

- appearance  
  Specifies the appearance of the content control.

- color  
  Specifies the red-green-blue (RGB) value of the color of the content control. You can provide the value in the '#RRGGBB' format.

- context  
  The request context associated with the object. This connects the add-in's process to the Office host application's process.

- dateCalendarType  
  Specifies a CalendarType value that represents the calendar type for the date picker content control.

- dateDisplayFormat  
  Specifies the format in which dates are displayed.

- dateDisplayLocale  
  Specifies a LanguageId that represents the language format for the date displayed in the date picker content control.

- dateStorageFormat  
  Specifies a ContentControlDateStorageFormat value that represents the format for storage and retrieval of dates when the date picker content control is bound to the XML data store of the active document.

- id  
  Gets the identification for the content control.

- isTemporary  
  Specifies whether to remove the content control from the active document when the user edits the contents of the control.

- level  
  Specifies the level of the content controlâwhether the content control surrounds text, paragraphs, table cells, or table rows; or if it is inline.

- lockContentControl  
  Specifies if the content control is locked (can't be deleted). true means that the user can't delete it from the active document, false means it can be deleted.

- lockContents  
  Specifies if the contents of the content control are locked (not editable). true means the user can't edit the contents, false means the contents are editable.

- placeholderText  
  Returns a BuildingBlock object that represents the placeholder text for the content control.

- range  
  Gets a Range object that represents the contents of the content control in the active document.

- showingPlaceholderText  
  Gets whether the placeholder text for the content control is being displayed.

- tag  
  Specifies a tag to identify the content control.

- title  
  Specifies the title for the content control.

- xmlMapping  
  Gets an XmlMapping object that represents the mapping of the content control to XML data in the data store of the document.

## Methods

- copy()  
  Copies the content control from the active document to the Clipboard.

- cut()  
  Removes the content control from the active document and moves the content control to the Clipboard.

- delete(deleteContents)  
  Deletes this content control and the contents of the content control.

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

- setPlaceholderText(options)  
  Sets the placeholder text that displays in the content control until a user enters their own text.

- toJSON()  
  Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). Whereas the original Word.DatePickerContentControl object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.DatePickerContentControlData) that contains shallow copies of any loaded child properties from the original object.

- track()  
  Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

- untrack()  
  Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

## Property Details

### appearance

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the appearance of the content control.

```typescript
appearance: Word.ContentControlAppearance | "BoundingBox" | "Tags" | "Hidden";
```

Property Value: Word.ContentControlAppearance | "BoundingBox" | "Tags" | "Hidden"

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

### color

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the red-green-blue (RGB) value of the color of the content control. You can provide the value in the '#RRGGBB' format.

```typescript
color: string;
```

Property Value: string

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

### context

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value: Word.RequestContext

### dateCalendarType

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a CalendarType value that represents the calendar type for the date picker content control.

```typescript
dateCalendarType: Word.CalendarType | "Western" | "Arabic" | "Hebrew" | "Taiwan" | "Japan" | "Thai" | "Korean" | "SakaEra" | "TranslitEnglish" | "TranslitFrench" | "Umalqura";
```

Property Value: Word.CalendarType | "Western" | "Arabic" | "Hebrew" | "Taiwan" | "Japan" | "Thai" | "Korean" | "SakaEra" | "TranslitEnglish" | "TranslitFrench" | "Umalqura"

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

### dateDisplayFormat

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the format in which dates are displayed.

```typescript
dateDisplayFormat: string;
```

Property Value: string

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

### dateDisplayLocale

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a LanguageId that represents the language format for the date displayed in the date picker content control.

```typescript
dateDisplayLocale: Word.LanguageId | "Afrikaans" | "Albanian" | "Amharic" | "Arabic" | "ArabicAlgeria" | "ArabicBahrain" | "ArabicEgypt" | "ArabicIraq" | "ArabicJordan" | "ArabicKuwait" | "ArabicLebanon" | "ArabicLibya" | "ArabicMorocco" | "ArabicOman" | "ArabicQatar" | "ArabicSyria" | "ArabicTunisia" | "ArabicUAE" | "ArabicYemen" | "Armenian" | "Assamese" | "AzeriCyrillic" | "AzeriLatin" | "Basque" | "BelgianDutch" | "BelgianFrench" | "Bengali" | "Bulgarian" | "Burmese" | "Belarusian" | "Catalan" | "Cherokee" | "ChineseHongKongSAR" | "ChineseMacaoSAR" | "ChineseSingapore" | "Croatian" | "Czech" | "Danish" | "Divehi" | "Dutch" | "Edo" | "EnglishAUS" | "EnglishBelize" | "EnglishCanadian" | "EnglishCaribbean" | "EnglishIndonesia" | "EnglishIreland" | "EnglishJamaica" | "EnglishNewZealand" | "EnglishPhilippines" | "EnglishSouthAfrica" | "EnglishTrinidadTobago" | "EnglishUK" | "EnglishUS" | "EnglishZimbabwe" | "Estonian" | "Faeroese" | "Filipino" | "Finnish" | "French" | "FrenchCameroon" | "FrenchCanadian" | "FrenchCongoDRC" | "FrenchCotedIvoire" | "FrenchHaiti" | "FrenchLuxembourg" | "FrenchMali" | "FrenchMonaco" | "FrenchMorocco" | "FrenchReunion" | "FrenchSenegal" | "FrenchWestIndies" | "FrisianNetherlands" | "Fulfulde" | "GaelicIreland" | "GaelicScotland" | "Galician" | "Georgian" | "German" | "GermanAustria" | "GermanLiechtenstein" | "GermanLuxembourg" | "Greek" | "Guarani" | "Gujarati" | "Hausa" | "Hawaiian" | "Hebrew" | "Hindi" | "Hungarian" | "Ibibio" | "Icelandic" | "Igbo" | "Indonesian" | "Inuktitut" | "Italian" | "Japanese" | "Kannada" | "Kanuri" | "Kashmiri" | "Kazakh" | "Khmer" | "Kirghiz" | "Konkani" | "Korean" | "Kyrgyz" | "LanguageNone" | "Lao" | "Latin" | "Latvian" | "Lithuanian" | "MacedonianFYROM" | "Malayalam" | "MalayBruneiDarussalam" | "Malaysian" | "Maltese" | "Manipuri" | "Marathi" | "MexicanSpanish" | "Mongolian" | "Nepali" | "NoProofing" | "NorwegianBokmol" | "NorwegianNynorsk" | "Oriya" | "Oromo" | "Pashto" | "Persian" | "Polish" | "Portuguese" | "PortugueseBrazil" | "Punjabi" | "RhaetoRomanic" | "Romanian" | "RomanianMoldova" | "Russian" | "RussianMoldova" | "SamiLappish" | "Sanskrit" | "SerbianCyrillic" | "SerbianLatin" | "Sesotho" | "SimplifiedChinese" | "Sindhi" | "SindhiPakistan" | "Sinhalese" | "Slovak" | "Slovenian" | "Somali" | "Sorbian" | "Spanish" | "SpanishArgentina" | "SpanishBolivia" | "SpanishChile" | "SpanishColombia" | "SpanishCostaRica" | "SpanishDominicanRepublic" | "SpanishEcuador" | "SpanishElSalvador" | "SpanishGuatemala" | "SpanishHonduras" | "SpanishModernSort" | "SpanishNicaragua" | "SpanishPanama" | "SpanishParaguay" | "SpanishPeru" | "SpanishPuertoRico" | "SpanishUruguay" | "SpanishVenezuela" | "Sutu" | "Swahili" | "Swedish" | "SwedishFinland" | "SwissFrench" | "SwissGerman" | "SwissItalian" | "Syriac" | "Tajik" | "Tamazight" | "TamazightLatin" | "Tamil" | "Tatar" | "Telugu" | "Thai" | "Tibetan" | "TigrignaEritrea" | "TigrignaEthiopic" | "TraditionalChinese" | "Tsonga" | "Tswana" | "Turkish" | "Turkmen" | "Ukrainian" | "Urdu" | "UzbekCyrillic" | "UzbekLatin" | "Venda" | "Vietnamese" | "Welsh" | "Xhosa" | "Yi" | "Yiddish" | "Yoruba" | "Zulu";
```

Property Value: Word.LanguageId | "Afrikaans" | "Albanian" | "Amharic" | "Arabic" | "ArabicAlgeria" | "ArabicBahrain" | "ArabicEgypt" | "ArabicIraq" | "ArabicJordan" | "ArabicKuwait" | "ArabicLebanon" | "ArabicLibya" | "ArabicMorocco" | "ArabicOman" | "ArabicQatar" | "ArabicSyria" | "ArabicTunisia" | "ArabicUAE" | "ArabicYemen" | "Armenian" | "Assamese" | "AzeriCyrillic" | "AzeriLatin" | "Basque" | "BelgianDutch" | "BelgianFrench" | "Bengali" | "Bulgarian" | "Burmese" | "Belarusian" | "Catalan" | "Cherokee" | "ChineseHongKongSAR" | "ChineseMacaoSAR" | "ChineseSingapore" | "Croatian" | "Czech" | "Danish" | "Divehi" | "Dutch" | "Edo" | "EnglishAUS" | "EnglishBelize" | "EnglishCanadian" | "EnglishCaribbean" | "EnglishIndonesia" | "EnglishIreland" | "EnglishJamaica" | "EnglishNewZealand" | "EnglishPhilippines" | "EnglishSouthAfrica" | "EnglishTrinidadTobago" | "EnglishUK" | "EnglishUS" | "EnglishZimbabwe" | "Estonian" | "Faeroese" | "Filipino" | "Finnish" | "French" | "FrenchCameroon" | "FrenchCanadian" | "FrenchCongoDRC" | "FrenchCotedIvoire" | "FrenchHaiti" | "FrenchLuxembourg" | "FrenchMali" | "FrenchMonaco" | "FrenchMorocco" | "FrenchReunion" | "FrenchSenegal" | "FrenchWestIndies" | "FrisianNetherlands" | "Fulfulde" | "GaelicIreland" | "GaelicScotland" | "Galician" | "Georgian" | "German" | "GermanAustria" | "GermanLiechtenstein" | "GermanLuxembourg" | "Greek" | "Guarani" | "Gujarati" | "Hausa" | "Hawaiian" | "Hebrew" | "Hindi" | "Hungarian" | "Ibibio" | "Icelandic" | "Igbo" | "Indonesian" | "Inuktitut" | "Italian" | "Japanese" | "Kannada" | "Kanuri" | "Kashmiri" | "Kazakh" | "Khmer" | "Kirghiz" | "Konkani" | "Korean" | "Kyrgyz" | "LanguageNone" | "Lao" | "Latin" | "Latvian" | "Lithuanian" | "MacedonianFYROM" | "Malayalam" | "MalayBruneiDarussalam" | "Malaysian" | "Maltese" | "Manipuri" | "Marathi" | "MexicanSpanish" | "Mongolian" | "Nepali" | "NoProofing" | "NorwegianBokmol" | "NorwegianNynorsk" | "Oriya" | "Oromo" | "Pashto" | "Persian" | "Polish" | "Portuguese" | "PortugueseBrazil" | "Punjabi" | "RhaetoRomanic" | "Romanian" | "RomanianMoldova" | "Russian" | "RussianMoldova" | "SamiLappish" | "Sanskrit" | "SerbianCyrillic" | "SerbianLatin" | "Sesotho" | "SimplifiedChinese" | "Sindhi" | "SindhiPakistan" | "Sinhalese" | "Slovak" | "Slovenian" | "Somali" | "Sorbian" | "Spanish" | "SpanishArgentina" | "SpanishBolivia" | "SpanishChile" | "SpanishColombia" | "SpanishCostaRica" | "SpanishDominicanRepublic" | "SpanishEcuador" | "SpanishElSalvador" | "SpanishGuatemala" | "SpanishHonduras" | "SpanishModernSort" | "SpanishNicaragua" | "SpanishPanama" | "SpanishParaguay" | "SpanishPeru" | "SpanishPuertoRico" | "SpanishUruguay" | "SpanishVenezuela" | "Sutu" | "Swahili" | "Swedish" | "SwedishFinland" | "SwissFrench" | "SwissGerman" | "SwissItalian" | "Syriac" | "Tajik" | "Tamazight" | "TamazightLatin" | "Tamil" | "Tatar" | "Telugu" | "Thai" | "Tibetan" | "TigrignaEritrea" | "TigrignaEthiopic" | "TraditionalChinese" | "Tsonga" | "Tswana" | "Turkish" | "Turkmen" | "Ukrainian" | "Urdu" | "UzbekCyrillic" | "UzbekLatin" | "Venda" | "Vietnamese" | "Welsh" | "Xhosa" | "Yi" | "Yiddish" | "Yoruba" | "Zulu"

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

### dateStorageFormat

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a ContentControlDateStorageFormat value that represents the format for storage and retrieval of dates when the date picker content control is bound to the XML data store of the active document.

```typescript
dateStorageFormat: Word.ContentControlDateStorageFormat | "Text" | "Date" | "DateTime";
```

Property Value: Word.ContentControlDateStorageFormat | "Text" | "Date" | "DateTime"

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

### id

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the identification for the content control.

```typescript
readonly id: string;
```

Property Value: string

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

### isTemporary

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether to remove the content control from the active document when the user edits the contents of the control.

```typescript
isTemporary: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

### level

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the level of the content controlâwhether the content control surrounds text, paragraphs, table cells, or table rows; or if it is inline.

```typescript
readonly level: Word.ContentControlLevel | "Inline" | "Paragraph" | "Row" | "Cell";
```

Property Value: Word.ContentControlLevel | "Inline" | "Paragraph" | "Row" | "Cell"

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

### lockContentControl

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the content control is locked (can't be deleted). true means that the user can't delete it from the active document, false means it can be deleted.

```typescript
lockContentControl: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

### lockContents

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the contents of the content control are locked (not editable). true means the user can't edit the contents, false means the contents are editable.

```typescript
lockContents: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

### placeholderText

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a BuildingBlock object that represents the placeholder text for the content control.

```typescript
readonly placeholderText: Word.BuildingBlock;
```

Property Value: Word.BuildingBlock

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

### range

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets a Range object that represents the contents of the content control in the active document.

```typescript
readonly range: Word.Range;
```

Property Value: Word.Range

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

### showingPlaceholderText

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether the placeholder text for the content control is being displayed.

```typescript
readonly showingPlaceholderText: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

### tag

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a tag to identify the content control.

```typescript
tag: string;
```

Property Value: string

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

### title

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the title for the content control.

```typescript
title: string;
```

Property Value: string

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

### xmlMapping

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets an XmlMapping object that represents the mapping of the content control to XML data in the data store of the document.

```typescript
readonly xmlMapping: Word.XmlMapping;
```

Property Value: Word.XmlMapping

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

## Method Details

### copy()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Copies the content control from the active document to the Clipboard.

```typescript
copy(): void;
```

Returns: void

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

### cut()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Removes the content control from the active document and moves the content control to the Clipboard.

```typescript
cut(): void;
```

Returns: void

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

### delete(deleteContents)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Deletes this content control and the contents of the content control.

```typescript
delete(deleteContents?: boolean): void;
```

Parameters:
- deleteContents: boolean

Optional. If true, deletes the contents as well.

Returns: void

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

### load(options)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(options?: Word.Interfaces.DatePickerContentControlLoadOptions): Word.DatePickerContentControl;
```

Parameters:
- options: Word.Interfaces.DatePickerContentControlLoadOptions

Provides options for which properties of the object to load.

Returns: Word.DatePickerContentControl

### load(propertyNames)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.DatePickerContentControl;
```

Parameters:
- propertyNames: string | string[]

A comma-delimited string or an array of strings that specify the properties to load.

Returns: Word.DatePickerContentControl

### load(propertyNamesAndPaths)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.DatePickerContentControl;
```

Parameters:
- propertyNamesAndPaths: { select?: string; expand?: string; }

propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

Returns: Word.DatePickerContentControl

### set(properties, options)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.DatePickerContentControlUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

Parameters:
- properties: Word.Interfaces.DatePickerContentControlUpdateData

A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
- options: OfficeExtension.UpdateOptions

Provides an option to suppress errors if the properties object tries to set any read-only properties.

Returns: void

### set(properties)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.DatePickerContentControl): void;
```

Parameters:
- properties: Word.DatePickerContentControl

Returns: void

### setPlaceholderText(options)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets the placeholder text that displays in the content control until a user enters their own text.

```typescript
setPlaceholderText(options?: Word.ContentControlPlaceholderOptions): void;
```

Parameters:
- options: Word.ContentControlPlaceholderOptions

Optional. The options for configuring the content control's placeholder text.

Returns: void

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

### toJSON()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.DatePickerContentControl object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.DatePickerContentControlData) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.DatePickerContentControlData;
```

Returns: Word.Interfaces.DatePickerContentControlData

### track()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.DatePickerContentControl;
```

Returns: Word.DatePickerContentControl

### untrack()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.DatePickerContentControl;
```

Returns: Word.DatePickerContentControl