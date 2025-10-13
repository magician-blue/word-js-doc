# Word.Index class

Package: [word](/en-us/javascript/api/word)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents a single index. The Index object is a member of the [Word.IndexCollection](/en-us/javascript/api/word/word.indexcollection). The IndexCollection includes all the indexes in the document.

Extends
- [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

## Properties
- [context](#word-word-index-context-member)  
  The request context associated with the object. This connects the add-in's process to the Office host application's process.

- [filter](#word-word-index-filter-member)  
  Gets a value that represents how Microsoft Word classifies the first character of entries in the index. See IndexFilter for available values.

- [headingSeparator](#word-word-index-headingseparator-member)  
  Gets the text between alphabetical groups (entries that start with the same letter) in the index. Corresponds to the \h switch for an [INDEX field](https://support.microsoft.com/office/adafcf4a-cb30-43f6-85c7-743da1635d9e).

- [indexLanguage](#word-word-index-indexlanguage-member)  
  Gets a LanguageId value that represents the sorting language to use for the index.

- [numberOfColumns](#word-word-index-numberofcolumns-member)  
  Gets the number of columns for each page of the index.

- [range](#word-word-index-range-member)  
  Returns a Range object that represents the portion of the document that is contained within the index.

- [rightAlignPageNumbers](#word-word-index-rightalignpagenumbers-member)  
  Specifies if page numbers are aligned with the right margin in the index.

- [separateAccentedLetterHeadings](#word-word-index-separateaccentedletterheadings-member)  
  Gets if the index contains separate headings for accented letters (for example, words that begin with "Ã" are under one heading and words that begin with "A" are under another).

- [sortBy](#word-word-index-sortby-member)  
  Specifies the sorting criteria for the index.

- [tabLeader](#word-word-index-tableader-member)  
  Specifies the leader character between entries in the index and their associated page numbers.

- [type](#word-word-index-type-member)  
  Gets the index type.

## Methods
- [delete()](#word-word-index-delete-member1)  
  Deletes this index.

- [load(options)](#word-word-index-load-member1)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

- [load(propertyNames)](#word-word-index-load-member2)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

- [load(propertyNamesAndPaths)](#word-word-index-load-member3)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

- [set(properties, options)](#word-word-index-set-member1)  
  Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

- [set(properties)](#word-word-index-set-member2)  
  Sets multiple properties on the object at the same time, based on an existing loaded object.

- [toJSON()](#word-word-index-tojson-member1)  
  Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.Index object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.IndexData) that contains shallow copies of any loaded child properties from the original object.

- [track()](#word-word-index-track-member1)  
  Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

- [untrack()](#word-word-index-untrack-member1)  
  Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

## Property Details

<a id="word-word-index-context-member"></a>
### context

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value
- [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

<a id="word-word-index-filter-member"></a>
### filter

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets a value that represents how Microsoft Word classifies the first character of entries in the index. See IndexFilter for available values.

```typescript
readonly filter: Word.IndexFilter | "None" | "Aiueo" | "Akasatana" | "Chosung" | "Low" | "Medium" | "Full";
```

Property Value
- [Word.IndexFilter](/en-us/javascript/api/word/word.indexfilter) | "None" | "Aiueo" | "Akasatana" | "Chosung" | "Low" | "Medium" | "Full"

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

<a id="word-word-index-headingseparator-member"></a>
### headingSeparator

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the text between alphabetical groups (entries that start with the same letter) in the index. Corresponds to the \h switch for an [INDEX field](https://support.microsoft.com/office/adafcf4a-cb30-43f6-85c7-743da1635d9e).

```typescript
readonly headingSeparator: Word.HeadingSeparator | "None" | "BlankLine" | "Letter" | "LetterLow" | "LetterFull";
```

Property Value
- [Word.HeadingSeparator](/en-us/javascript/api/word/word.headingseparator) | "None" | "BlankLine" | "Letter" | "LetterLow" | "LetterFull"

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

<a id="word-word-index-indexlanguage-member"></a>
### indexLanguage

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets a LanguageId value that represents the sorting language to use for the index.

```typescript
readonly indexLanguage: Word.LanguageId | "Afrikaans" | "Albanian" | "Amharic" | "Arabic" | "ArabicAlgeria" | "ArabicBahrain" | "ArabicEgypt" | "ArabicIraq" | "ArabicJordan" | "ArabicKuwait" | "ArabicLebanon" | "ArabicLibya" | "ArabicMorocco" | "ArabicOman" | "ArabicQatar" | "ArabicSyria" | "ArabicTunisia" | "ArabicUAE" | "ArabicYemen" | "Armenian" | "Assamese" | "AzeriCyrillic" | "AzeriLatin" | "Basque" | "BelgianDutch" | "BelgianFrench" | "Bengali" | "Bulgarian" | "Burmese" | "Belarusian" | "Catalan" | "Cherokee" | "ChineseHongKongSAR" | "ChineseMacaoSAR" | "ChineseSingapore" | "Croatian" | "Czech" | "Danish" | "Divehi" | "Dutch" | "Edo" | "EnglishAUS" | "EnglishBelize" | "EnglishCanadian" | "EnglishCaribbean" | "EnglishIndonesia" | "EnglishIreland" | "EnglishJamaica" | "EnglishNewZealand" | "EnglishPhilippines" | "EnglishSouthAfrica" | "EnglishTrinidadTobago" | "EnglishUK" | "EnglishUS" | "EnglishZimbabwe" | "Estonian" | "Faeroese" | "Filipino" | "Finnish" | "French" | "FrenchCameroon" | "FrenchCanadian" | "FrenchCongoDRC" | "FrenchCotedIvoire" | "FrenchHaiti" | "FrenchLuxembourg" | "FrenchMali" | "FrenchMonaco" | "FrenchMorocco" | "FrenchReunion" | "FrenchSenegal" | "FrenchWestIndies" | "FrisianNetherlands" | "Fulfulde" | "GaelicIreland" | "GaelicScotland" | "Galician" | "Georgian" | "German" | "GermanAustria" | "GermanLiechtenstein" | "GermanLuxembourg" | "Greek" | "Guarani" | "Gujarati" | "Hausa" | "Hawaiian" | "Hebrew" | "Hindi" | "Hungarian" | "Ibibio" | "Icelandic" | "Igbo" | "Indonesian" | "Inuktitut" | "Italian" | "Japanese" | "Kannada" | "Kanuri" | "Kashmiri" | "Kazakh" | "Khmer" | "Kirghiz" | "Konkani" | "Korean" | "Kyrgyz" | "LanguageNone" | "Lao" | "Latin" | "Latvian" | "Lithuanian" | "MacedonianFYROM" | "Malayalam" | "MalayBruneiDarussalam" | "Malaysian" | "Maltese" | "Manipuri" | "Marathi" | "MexicanSpanish" | "Mongolian" | "Nepali" | "NoProofing" | "NorwegianBokmol" | "NorwegianNynorsk" | "Oriya" | "Oromo" | "Pashto" | "Persian" | "Polish" | "Portuguese" | "PortugueseBrazil" | "Punjabi" | "RhaetoRomanic" | "Romanian" | "RomanianMoldova" | "Russian" | "RussianMoldova" | "SamiLappish" | "Sanskrit" | "SerbianCyrillic" | "SerbianLatin" | "Sesotho" | "SimplifiedChinese" | "Sindhi" | "SindhiPakistan" | "Sinhalese" | "Slovak" | "Slovenian" | "Somali" | "Sorbian" | "Spanish" | "SpanishArgentina" | "SpanishBolivia" | "SpanishChile" | "SpanishColombia" | "SpanishCostaRica" | "SpanishDominicanRepublic" | "SpanishEcuador" | "SpanishElSalvador" | "SpanishGuatemala" | "SpanishHonduras" | "SpanishModernSort" | "SpanishNicaragua" | "SpanishPanama" | "SpanishParaguay" | "SpanishPeru" | "SpanishPuertoRico" | "SpanishUruguay" | "SpanishVenezuela" | "Sutu" | "Swahili" | "Swedish" | "SwedishFinland" | "SwissFrench" | "SwissGerman" | "SwissItalian" | "Syriac" | "Tajik" | "Tamazight" | "TamazightLatin" | "Tamil" | "Tatar" | "Telugu" | "Thai" | "Tibetan" | "TigrignaEritrea" | "TigrignaEthiopic" | "TraditionalChinese" | "Tsonga" | "Tswana" | "Turkish" | "Turkmen" | "Ukrainian" | "Urdu" | "UzbekCyrillic" | "UzbekLatin" | "Venda" | "Vietnamese" | "Welsh" | "Xhosa" | "Yi" | "Yiddish" | "Yoruba" | "Zulu";
```

Property Value
- [Word.LanguageId](/en-us/javascript/api/word/word.languageid) | "Afrikaans" | "Albanian" | "Amharic" | "Arabic" | "ArabicAlgeria" | "ArabicBahrain" | "ArabicEgypt" | "ArabicIraq" | "ArabicJordan" | "ArabicKuwait" | "ArabicLebanon" | "ArabicLibya" | "ArabicMorocco" | "ArabicOman" | "ArabicQatar" | "ArabicSyria" | "ArabicTunisia" | "ArabicUAE" | "ArabicYemen" | "Armenian" | "Assamese" | "AzeriCyrillic" | "AzeriLatin" | "Basque" | "BelgianDutch" | "BelgianFrench" | "Bengali" | "Bulgarian" | "Burmese" | "Belarusian" | "Catalan" | "Cherokee" | "ChineseHongKongSAR" | "ChineseMacaoSAR" | "ChineseSingapore" | "Croatian" | "Czech" | "Danish" | "Divehi" | "Dutch" | "Edo" | "EnglishAUS" | "EnglishBelize" | "EnglishCanadian" | "EnglishCaribbean" | "EnglishIndonesia" | "EnglishIreland" | "EnglishJamaica" | "EnglishNewZealand" | "EnglishPhilippines" | "EnglishSouthAfrica" | "EnglishTrinidadTobago" | "EnglishUK" | "EnglishUS" | "EnglishZimbabwe" | "Estonian" | "Faeroese" | "Filipino" | "Finnish" | "French" | "FrenchCameroon" | "FrenchCanadian" | "FrenchCongoDRC" | "FrenchCotedIvoire" | "FrenchHaiti" | "FrenchLuxembourg" | "FrenchMali" | "FrenchMonaco" | "FrenchMorocco" | "FrenchReunion" | "FrenchSenegal" | "FrenchWestIndies" | "FrisianNetherlands" | "Fulfulde" | "GaelicIreland" | "GaelicScotland" | "Galician" | "Georgian" | "German" | "GermanAustria" | "GermanLiechtenstein" | "GermanLuxembourg" | "Greek" | "Guarani" | "Gujarati" | "Hausa" | "Hawaiian" | "Hebrew" | "Hindi" | "Hungarian" | "Ibibio" | "Icelandic" | "Igbo" | "Indonesian" | "Inuktitut" | "Italian" | "Japanese" | "Kannada" | "Kanuri" | "Kashmiri" | "Kazakh" | "Khmer" | "Kirghiz" | "Konkani" | "Korean" | "Kyrgyz" | "LanguageNone" | "Lao" | "Latin" | "Latvian" | "Lithuanian" | "MacedonianFYROM" | "Malayalam" | "MalayBruneiDarussalam" | "Malaysian" | "Maltese" | "Manipuri" | "Marathi" | "MexicanSpanish" | "Mongolian" | "Nepali" | "NoProofing" | "NorwegianBokmol" | "NorwegianNynorsk" | "Oriya" | "Oromo" | "Pashto" | "Persian" | "Polish" | "Portuguese" | "PortugueseBrazil" | "Punjabi" | "RhaetoRomanic" | "Romanian" | "RomanianMoldova" | "Russian" | "RussianMoldova" | "SamiLappish" | "Sanskrit" | "SerbianCyrillic" | "SerbianLatin" | "Sesotho" | "SimplifiedChinese" | "Sindhi" | "SindhiPakistan" | "Sinhalese" | "Slovak" | "Slovenian" | "Somali" | "Sorbian" | "Spanish" | "SpanishArgentina" | "SpanishBolivia" | "SpanishChile" | "SpanishColombia" | "SpanishCostaRica" | "SpanishDominicanRepublic" | "SpanishEcuador" | "SpanishElSalvador" | "SpanishGuatemala" | "SpanishHonduras" | "SpanishModernSort" | "SpanishNicaragua" | "SpanishPanama" | "SpanishParaguay" | "SpanishPeru" | "SpanishPuertoRico" | "SpanishUruguay" | "SpanishVenezuela" | "Sutu" | "Swahili" | "Swedish" | "SwedishFinland" | "SwissFrench" | "SwissGerman" | "SwissItalian" | "Syriac" | "Tajik" | "Tamazight" | "TamazightLatin" | "Tamil" | "Tatar" | "Telugu" | "Thai" | "Tibetan" | "TigrignaEritrea" | "TigrignaEthiopic" | "TraditionalChinese" | "Tsonga" | "Tswana" | "Turkish" | "Turkmen" | "Ukrainian" | "Urdu" | "UzbekCyrillic" | "UzbekLatin" | "Venda" | "Vietnamese" | "Welsh" | "Xhosa" | "Yi" | "Yiddish" | "Yoruba" | "Zulu"

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

<a id="word-word-index-numberofcolumns-member"></a>
### numberOfColumns

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the number of columns for each page of the index.

```typescript
readonly numberOfColumns: number;
```

Property Value
- number

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

<a id="word-word-index-range-member"></a>
### range

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a Range object that represents the portion of the document that is contained within the index.

```typescript
readonly range: Word.Range;
```

Property Value
- [Word.Range](/en-us/javascript/api/word/word.range)

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

<a id="word-word-index-rightalignpagenumbers-member"></a>
### rightAlignPageNumbers

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if page numbers are aligned with the right margin in the index.

```typescript
readonly rightAlignPageNumbers: boolean;
```

Property Value
- boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

<a id="word-word-index-separateaccentedletterheadings-member"></a>
### separateAccentedLetterHeadings

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets if the index contains separate headings for accented letters (for example, words that begin with "Ã" are under one heading and words that begin with "A" are under another).

```typescript
readonly separateAccentedLetterHeadings: boolean;
```

Property Value
- boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

<a id="word-word-index-sortby-member"></a>
### sortBy

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the sorting criteria for the index.

```typescript
readonly sortBy: Word.IndexSortBy | "Stroke" | "Syllable";
```

Property Value
- [Word.IndexSortBy](/en-us/javascript/api/word/word.indexsortby) | "Stroke" | "Syllable"

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

<a id="word-word-index-tableader-member"></a>
### tabLeader

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the leader character between entries in the index and their associated page numbers.

```typescript
tabLeader: Word.TabLeader | "Spaces" | "Dots" | "Dashes" | "Lines" | "Heavy" | "MiddleDot";
```

Property Value
- [Word.TabLeader](/en-us/javascript/api/word/word.tableader) | "Spaces" | "Dots" | "Dashes" | "Lines" | "Heavy" | "MiddleDot"

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

<a id="word-word-index-type-member"></a>
### type

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the index type.

```typescript
readonly type: Word.IndexType | "Indent" | "Runin";
```

Property Value
- [Word.IndexType](/en-us/javascript/api/word/word.indextype) | "Indent" | "Runin"

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

## Method Details

<a id="word-word-index-delete-member1"></a>
### delete()

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Deletes this index.

```typescript
delete(): void;
```

Returns
- void

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

<a id="word-word-index-load-member1"></a>
### load(options)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(options?: Word.Interfaces.IndexLoadOptions): Word.Index;
```

Parameters
- options: [Word.Interfaces.IndexLoadOptions](/en-us/javascript/api/word/word.interfaces.indexloadoptions)  
  Provides options for which properties of the object to load.

Returns
- [Word.Index](/en-us/javascript/api/word/word.index)

<a id="word-word-index-load-member2"></a>
### load(propertyNames)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.Index;
```

Parameters
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns
- [Word.Index](/en-us/javascript/api/word/word.index)

<a id="word-word-index-load-member3"></a>
### load(propertyNamesAndPaths)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
    select?: string;
    expand?: string;
}): Word.Index;
```

Parameters
- propertyNamesAndPaths: { select?: string; expand?: string; }  
  `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

Returns
- [Word.Index](/en-us/javascript/api/word/word.index)

<a id="word-word-index-set-member1"></a>
### set(properties, options)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.IndexUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

Parameters
- properties: [Word.Interfaces.IndexUpdateData](/en-us/javascript/api/word/word.interfaces.indexupdatedata)  
  A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
- options: [OfficeExtension.UpdateOptions](/en-us/javascript/api/office/officeextension.updateoptions)  
  Provides an option to suppress errors if the properties object tries to set any read-only properties.

Returns
- void

<a id="word-word-index-set-member2"></a>
### set(properties)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.Index): void;
```

Parameters
- properties: [Word.Index](/en-us/javascript/api/word/word.index)

Returns
- void

<a id="word-word-index-tojson-member1"></a>
### toJSON()

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.Index` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.IndexData`) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.IndexData;
```

Returns
- [Word.Interfaces.IndexData](/en-us/javascript/api/word/word.interfaces.indexdata)

<a id="word-word-index-track-member1"></a>
### track()

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.Index;
```

Returns
- [Word.Index](/en-us/javascript/api/word/word.index)

<a id="word-word-index-untrack-member1"></a>
### untrack()

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

```typescript
untrack(): Word.Index;
```

Returns
- [Word.Index](/en-us/javascript/api/word/word.index)