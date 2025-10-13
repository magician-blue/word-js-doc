# Word.FieldType enum

Package: [word](/en-us/javascript/api/word)

Represents the type of Field.

## Remarks

[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples

```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-fields.yaml

// Inserts a Date field before selection.
await Word.run(async (context) => {
  const range: Word.Range = context.document.getSelection().getRange();

  const field: Word.Field = range.insertField(Word.InsertLocation.before, Word.FieldType.date, '\\@ "M/d/yyyy h:mm am/pm"', true);

  field.load("result,code");
  await context.sync();

  if (field.isNullObject) {
    console.log("There are no fields in this document.");
  } else {
    console.log("Code of the field: " + field.code, "Result of the field: " + JSON.stringify(field.result));
  }
});
```

## Fields

- addin = "Addin"
  - Represents that the field type is Add-in.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- addressBlock = "AddressBlock"
  - Represents that the field type is AddressBlock.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- advance = "Advance"
  - Represents that the field type is Advance.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- ask = "Ask"
  - Represents that the field type is Ask.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- author = "Author"
  - Represents that the field type is Author.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- autoText = "AutoText"
  - Represents that the field type is AutoText.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- autoTextList = "AutoTextList"
  - Represents that the field type is AutoTextList.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- barCode = "BarCode"
  - Represents that the field type is Barcode.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- bibliography = "Bibliography"
  - Represents that the field type is Bibliography.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- bidiOutline = "BidiOutline"
  - Represents that the field type is BidiOutline.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- citation = "Citation"
  - Represents that the field type is Citation.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- comments = "Comments"
  - Represents that the field type is Comments.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- compare = "Compare"
  - Represents that the field type is Compare.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- createDate = "CreateDate"
  - Represents that the field type is CreateDate.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- data = "Data"
  - Represents that the field type is Data.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- database = "Database"
  - Represents that the field type is Database.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- date = "Date"
  - Represents that the field type is Date.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- displayBarcode = "DisplayBarcode"
  - Represents that the field type is DisplayBarcode.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- docProperty = "DocProperty"
  - Represents that the field type is DocumentProperty
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- docVariable = "DocVariable"
  - Represents that the field type is DocumentVariable.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- editTime = "EditTime"
  - Represents that the field type is EditTime.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- embedded = "Embedded"
  - Represents that the field type is Embedded.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- empty = "Empty"
  - Represents that the field type is Empty.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- eq = "EQ"
  - Represents that the field type is Equation.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- expression = "Expression"
  - Represents that the field type is Expression.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- fileName = "FileName"
  - Represents that the field type is FileName.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- fileSize = "FileSize"
  - Represents that the field type is FileSize.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- fillIn = "FillIn"
  - Represents that the field type is FillIn.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- formCheckbox = "FormCheckbox"
  - Represents that the field type is FormCheckbox.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- formDropdown = "FormDropdown"
  - Represents that the field type is FormDropdown.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- formText = "FormText"
  - Represents that the field type is FormText.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- gotoButton = "GotoButton"
  - Represents that the field type is GotoButton.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- greetingLine = "GreetingLine"
  - Represents that the field type is GreetingLine.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- hyperlink = "Hyperlink"
  - Represents that the field type is Hyperlink.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- if = "If"
  - Represents that the field type is If.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- import = "Import"
  - Represents that the field type is Import.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- include = "Include"
  - Represents that the field type is Include.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- includePicture = "IncludePicture"
  - Represents that the field type is IncludePicture.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- includeText = "IncludeText"
  - Represents that the field type is IncludeText.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- index = "Index"
  - Represents that the field type is Index.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- info = "Info"
  - Represents that the field type is Information.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- keywords = "Keywords"
  - Represents that the field type is Keywords.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- lastSavedBy = "LastSavedBy"
  - Represents that the field type is LastSavedBy.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- link = "Link"
  - Represents that the field type is Link.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- listNum = "ListNum"
  - Represents that the field type is ListNumber.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- macroButton = "MacroButton"
  - Represents that the field type is MacroButton.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- mergeBarcode = "MergeBarcode"
  - Represents that the field type is MergeBarcode.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- mergeField = "MergeField"
  - Represents that the field type is MergeField.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- mergeRec = "MergeRec"
  - Represents that the field type is MergeRecord.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- mergeSeq = "MergeSeq"
  - Represents that the field type is MergeSequence.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- next = "Next"
  - Represents that the field type is Next.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- nextIf = "NextIf"
  - Represents that the field type is NextIf.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- noteRef = "NoteRef"
  - Represents that the field type is NoteReference.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- numChars = "NumChars"
  - Represents that the field type is NumberOfCharacters.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- numPages = "NumPages"
  - Represents that the field type is NumberOfPages.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- numWords = "NumWords"
  - Represents that the field type is NumberOfWords.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- ocx = "OCX"
  - Represents that the field type is ActiveXControl.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- others = "Others"
  - Represents the field types not supported by the Office JavaScript API.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- page = "Page"
  - Represents that the field type is Page.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- pageRef = "PageRef"
  - Represents that the field type is PageReference.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- print = "Print"
  - Represents that the field type is Print.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- printDate = "PrintDate"
  - Represents that the field type is PrintDate.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- private = "Private"
  - Represents that the field type is Private.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- quote = "Quote"
  - Represents that the field type is Quote.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- rd = "RD"
  - Represents that the field type is ReferencedDocument.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- ref = "Ref"
  - Represents that the field type is Reference.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- revNum = "RevNum"
  - Represents that the field type is RevisionNumber.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- saveDate = "SaveDate"
  - Represents that the field type is SaveDate.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- section = "Section"
  - Represents that the field type is Section.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- sectionPages = "SectionPages"
  - Represents that the field type is SectionPages.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- seq = "Seq"
  - Represents that the field type is Sequence.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- set = "Set"
  - Represents that the field type is Set.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- shape = "Shape"
  - Represents that the field type is Shape.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- skipIf = "SkipIf"
  - Represents that the field type is SkipIf.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- styleRef = "StyleRef"
  - Represents that the field type is StyleReference.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- subject = "Subject"
  - Represents that the field type is Subject.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- subscriber = "Subscriber"
  - Represents that the field type is Subscriber.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- symbol = "Symbol"
  - Represents that the field type is Symbol.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- ta = "TA"
  - Represents that the field type is TableOfAuthoritiesEntry.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- tc = "TC"
  - Represents that the field type is TableOfContentsEntry.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- template = "Template"
  - Represents that the field type is Template.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- time = "Time"
  - Represents that the field type is Time.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- title = "Title"
  - Represents that the field type is Title.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- toa = "TOA"
  - Represents that the field type is TableOfAuthorities.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- toc = "TOC"
  - Represents that the field type is TableOfContents.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
  
  **Complete TOC Insertion Example:**
  ```typescript
  // Insert a complete table of contents with title
  await Word.run(async (context) => {
      const body = context.document.body;
      
      // 1. Insert title at document start
      const startRange = body.getRange(Word.RangeLocation.start);
      startRange.insertText("Table of Contents\n", Word.InsertLocation.start);
      
      // 2. Format title
      const paragraphs = body.paragraphs;
      paragraphs.load('items');
      await context.sync();
      
      const titleParagraph = paragraphs.items[0];
      titleParagraph.font.size = 14;
      titleParagraph.font.bold = true;
      
      // 3. Insert TOC field after title
      const afterTitle = startRange.getRange(Word.RangeLocation.end);
      const tocField = afterTitle.insertField(Word.InsertLocation.start, Word.FieldType.toc, 'TOC \\o "1-3" \\h \\z \\u', false);
      
      // 4. Add spacing after TOC
      afterTitle.insertText("\n", Word.InsertLocation.after);
      
      await context.sync();
      
      // 5. Update TOC field to display content
      tocField.updateResult();
      await context.sync();
  });
  ```
  
  **TOC Field Code Parameters:**
  - `\\o "1-3"` - Include outline levels 1-3 (headings 1-3)
  - `\\h` - Create hyperlinks for navigation
  - `\\z` - Hide page numbers
  - `\\u` - Use outline levels instead of styles
  
  **Important Notes:**
  - Use simple Range operations; avoid complex chained Range manipulations
  - Set title format through paragraphs collection for better reliability
  - Insert content sequentially rather than using multiple insertBreak calls
  - Always call updateResult() after inserting TOC field
- undefined = "Undefined"
  - Represents that the field type is Undefined.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- userAddress = "UserAddress"
  - Represents that the field type is UserAddress.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- userInitials = "UserInitials"
  - Represents that the field type is UserInitials.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- userName = "UserName"
  - Represents that the field type is UserName.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- xe = "XE"
  - Represents that the field type is IndexEntry.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)