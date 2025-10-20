## Interfaces

| Interface | Description |
|-----------|-------------|
| Word.AnnotationClickedEventArgs | Holds annotation information that is passed back on annotation inserted event. |
| Word.AnnotationHoveredEventArgs | Holds annotation information that is passed back on annotation hovered event. |
| Word.AnnotationInsertedEventArgs | Holds annotation information that is passed back on annotation added event. |
| Word.AnnotationPopupActionEventArgs | Represents action information that's passed back on annotation pop-up action event. |
| Word.AnnotationRemovedEventArgs | Holds annotation information that is passed back on annotation removed event. |
| Word.AnnotationSet | Annotations set produced by the add-in. Currently supporting only critiques. |
| Word.CommentDetail | A structure for the ID and reply IDs of this comment. |
| Word.CommentEventArgs | Provides information about the comments that raised the comment event. |
| Word.ContentControlAddedEventArgs | Provides information about the content control that raised contentControlAdded event. |
| Word.ContentControlDataChangedEventArgs | Provides information about the content control that raised contentControlDataChanged event. |
| Word.ContentControlDeletedEventArgs | Provides information about the content control that raised contentControlDeleted event. |
| Word.ContentControlEnteredEventArgs | Provides information about the content control that raised contentControlEntered event. |
| Word.ContentControlExitedEventArgs | Provides information about the content control that raised contentControlExited event. |
| Word.ContentControlOptions | Specifies the options that define which content controls are returned. |
| Word.ContentControlPlaceholderOptions | The options that define what placeholder to be used in the content control. |
| Word.ContentControlSelectionChangedEventArgs | Provides information about the content control that raised contentControlSelectionChanged event. |
| Word.Critique | Critique that will be rendered as underline for the specified part of paragraph in the document. |
| Word.CritiquePopupOptions | Properties defining the behavior of the pop-up menu for a given critique. |
| Word.CustomXmlAddNodeOptions | The options for adding a node to the XML tree. |
| Word.CustomXmlAddSchemaOptions | Adds one or more schemas to a schema collection that can then be added to a stream in the data store and to the schema library. |
| Word.CustomXmlAddValidationErrorOptions | The options that define the descriptive error text and the state of clearedOnUpdate. |
| Word.CustomXmlAppendChildNodeOptions | The options that define the prefix mapping and the source of the custom XML data. |
| Word.CustomXmlInsertNodeBeforeOptions | Inserts a new node just before the context node in the tree. |
| Word.CustomXmlInsertSubtreeBeforeOptions | Inserts a new node just before the context node in the tree. |
| Word.CustomXmlReplaceChildNodeOptions | Removes the specified child node and replaces it with a different node in the same location. |
| Word.DocumentCompareOptions | Specifies the options to be included in a compare document operation. |
| Word.GetTextOptions | Specifies the options to be included in a getText operation. |
| Word.HyperlinkAddOptions | Specifies the options for adding to a Word.HyperlinkCollection object. |
| Word.IndexAddOptions | Represents options for creating an index in a Word document. |
| Word.IndexMarkAllEntriesOptions | Represents options for marking all index entries in a Word document. |
| Word.IndexMarkEntryOptions | Represents options for marking an index entry in a Word document. |
| Word.InsertFileOptions | Specifies the options to determine what to copy when inserting a file. |
| Word.InsertShapeOptions | Specifies the options to determine location and size when inserting a shape. |
| Word.Interfaces.AnnotationCollectionData | An interface describing the data returned by calling annotationCollection.toJSON(). |
| Word.Interfaces.AnnotationCollectionLoadOptions | Contains a collection of Word.Annotation objects. |
| Word.Interfaces.AnnotationCollectionUpdateData | An interface for updating data on the AnnotationCollection object, for use in annotationCollection.set({ ... }). |
| Word.Interfaces.AnnotationData | An interface describing the data returned by calling annotation.toJSON(). |
| Word.Interfaces.AnnotationLoadOptions | Represents an annotation attached to a paragraph. |
| Word.Interfaces.ApplicationData | An interface describing the data returned by calling application.toJSON(). |
| Word.Interfaces.ApplicationLoadOptions | Represents the application object. |
| Word.Interfaces.ApplicationUpdateData | An interface for updating data on the Application object, for use in application.set({ ... }). |
| Word.Interfaces.BibliographyData | An interface describing the data returned by calling bibliography.toJSON(). |
| Word.Interfaces.BibliographyLoadOptions | Represents the list of available sources attached to the document (in the current list) or the list of sources available in the application (in the master list). |
| Word.Interfaces.BibliographyUpdateData | An interface for updating data on the Bibliography object, for use in bibliography.set({ ... }). |
| Word.Interfaces.BodyData | An interface describing the data returned by calling body.toJSON(). |
| Word.Interfaces.BodyLoadOptions | Represents the body of a document or a section. |
| Word.Interfaces.BodyUpdateData | An interface for updating data on the Body object, for use in body.set({ ... }). |
| Word.Interfaces.BookmarkCollectionData | An interface describing the data returned by calling bookmarkCollection.toJSON(). |
| Word.Interfaces.BookmarkCollectionLoadOptions | A collection of Word.Bookmark objects that represent the bookmarks in the specified selection, range, or document. |
| Word.Interfaces.BookmarkCollectionUpdateData | An interface for updating data on the BookmarkCollection object, for use in bookmarkCollection.set({ ... }). |
| Word.Interfaces.BookmarkData | An interface describing the data returned by calling bookmark.toJSON(). |
| Word.Interfaces.BookmarkLoadOptions | Represents a single bookmark in a document, selection, or range. The Bookmark object is a member of the Bookmark collection. The Word.BookmarkCollection includes all the bookmarks listed in the Bookmark dialog box (Insert menu). |
| Word.Interfaces.BookmarkUpdateData | An interface for updating data on the Bookmark object, for use in bookmark.set({ ... }). |
| Word.Interfaces.BorderCollectionData | An interface describing the data returned by calling borderCollection.toJSON(). |
| Word.Interfaces.BorderCollectionLoadOptions | Represents the collection of border styles. |
| Word.Interfaces.BorderCollectionUpdateData | An interface for updating data on the BorderCollection object, for use in borderCollection.set({ ... }). |
| Word.Interfaces.BorderData | An interface describing the data returned by calling border.toJSON(). |
| Word.Interfaces.BorderLoadOptions | Represents the Border object for text, a paragraph, or a table. |
| Word.Interfaces.BorderUniversalCollectionData | An interface describing the data returned by calling borderUniversalCollection.toJSON(). |
| Word.Interfaces.BorderUniversalCollectionLoadOptions | Represents the collection of Word.BorderUniversal objects. |
| Word.Interfaces.BorderUniversalCollectionUpdateData | An interface for updating data on the BorderUniversalCollection object, for use in borderUniversalCollection.set({ ... }). |
| Word.Interfaces.BorderUniversalData | An interface describing the data returned by calling borderUniversal.toJSON(). |
| Word.Interfaces.BorderUniversalLoadOptions | Represents the BorderUniversal object, which manages borders for a range, paragraph, table, or frame. |
| Word.Interfaces.BorderUniversalUpdateData | An interface for updating data on the BorderUniversal object, for use in borderUniversal.set({ ... }). |
| Word.Interfaces.BorderUpdateData | An interface for updating data on the Border object, for use in border.set({ ... }). |
| Word.Interfaces.BreakCollectionData | An interface describing the data returned by calling breakCollection.toJSON(). |
| Word.Interfaces.BreakCollectionLoadOptions | Contains a collection of Word.Break objects. |
| Word.Interfaces.BreakCollectionUpdateData | An interface for updating data on the BreakCollection object, for use in breakCollection.set({ ... }). |
| Word.Interfaces.BreakData | An interface describing the data returned by calling break.toJSON(). |
| Word.Interfaces.BreakLoadOptions | Represents a break in a Word document. |
| Word.Interfaces.BreakUpdateData | An interface for updating data on the Break object, for use in break.set({ ... }). |
| Word.Interfaces.BuildingBlockCategoryData | An interface describing the data returned by calling buildingBlockCategory.toJSON(). |
| Word.Interfaces.BuildingBlockCategoryLoadOptions | Represents a category of building blocks in a Word document. |
| Word.Interfaces.BuildingBlockData | An interface describing the data returned by calling buildingBlock.toJSON(). |
| Word.Interfaces.BuildingBlockGalleryContentControlData | An interface describing the data returned by calling buildingBlockGalleryContentControl.toJSON(). |
| Word.Interfaces.BuildingBlockGalleryContentControlLoadOptions | Represents the BuildingBlockGalleryContentControl object. |
| Word.Interfaces.BuildingBlockGalleryContentControlUpdateData | An interface for updating data on the BuildingBlockGalleryContentControl object, for use in buildingBlockGalleryContentControl.set({ ... }). |
| Word.Interfaces.BuildingBlockLoadOptions | Represents a building block in a template. A building block is pre-built content, similar to autotext, that may contain text, images, and formatting. |
| Word.Interfaces.BuildingBlockTypeItemData | An interface describing the data returned by calling buildingBlockTypeItem.toJSON(). |
| Word.Interfaces.BuildingBlockTypeItemLoadOptions | Represents a type of building block in a Word document. |
| Word.Interfaces.BuildingBlockUpdateData | An interface for updating data on the BuildingBlock object, for use in buildingBlock.set({ ... }). |
| Word.Interfaces.CanvasData | An interface describing the data returned by calling canvas.toJSON(). |
| Word.Interfaces.CanvasLoadOptions | Represents a canvas in the document. To get the corresponding Shape object, use Canvas.shape. |
| Word.Interfaces.CanvasUpdateData | An interface for updating data on the Canvas object, for use in canvas.set({ ... }). |
| Word.Interfaces.CheckboxContentControlData | An interface describing the data returned by calling checkboxContentControl.toJSON(). |
| Word.Interfaces.CheckboxContentControlLoadOptions | The data specific to content controls of type CheckBox. |
| Word.Interfaces.CheckboxContentControlUpdateData | An interface for updating data on the CheckboxContentControl object, for use in checkboxContentControl.set({ ... }). |
| Word.Interfaces.CollectionLoadOptions | Provides ways to load properties of only a subset of members of a collection. |
| Word.Interfaces.ColorFormatData | An interface describing the data returned by calling colorFormat.toJSON(). |
| Word.Interfaces.ColorFormatLoadOptions | Represents the color formatting of a shape or text in Word. |
| Word.Interfaces.ColorFormatUpdateData | An interface for updating data on the ColorFormat object, for use in colorFormat.set({ ... }). |
| Word.Interfaces.ComboBoxContentControlData | An interface describing the data returned by calling comboBoxContentControl.toJSON(). |
| Word.Interfaces.CommentCollectionData | An interface describing the data returned by calling commentCollection.toJSON(). |
| Word.Interfaces.CommentCollectionLoadOptions | Contains a collection of Word.Comment objects. |
| Word.Interfaces.CommentCollectionUpdateData | An interface for updating data on the CommentCollection object, for use in commentCollection.set({ ... }). |
| Word.Interfaces.CommentContentRangeData | An interface describing the data returned by calling commentContentRange.toJSON(). |
| Word.Interfaces.CommentContentRangeLoadOptions | |
| Word.Interfaces.CommentContentRangeUpdateData | An interface for updating data on the CommentContentRange object, for use in commentContentRange.set({ ... }). |
| Word.Interfaces.CommentData | An interface describing the data returned by calling comment.toJSON(). |
| Word.Interfaces.CommentLoadOptions | Represents a comment in the document. |
| Word.Interfaces.CommentReplyCollectionData | An interface describing the data returned by calling commentReplyCollection.toJSON(). |
| Word.Interfaces.CommentReplyCollectionLoadOptions | Contains a collection of Word.CommentReply objects. Represents all comment replies in one comment thread. |
| Word.Interfaces.CommentReplyCollectionUpdateData | An interface for updating data on the CommentReplyCollection object, for use in commentReplyCollection.set({ ... }). |
| Word.Interfaces.CommentReplyData | An interface describing the data returned by calling commentReply.toJSON(). |
| Word.Interfaces.CommentReplyLoadOptions | Represents a comment reply in the document. |
| Word.Interfaces.CommentReplyUpdateData | An interface for updating data on the CommentReply object, for use in commentReply.set({ ... }). |
| Word.Interfaces.CommentUpdateData | An interface for updating data on the Comment object, for use in comment.set({ ... }). |
| Word.Interfaces.ContentControlCollectionData | An interface describing the data returned by calling contentControlCollection.toJSON(). |
| Word.Interfaces.ContentControlCollectionLoadOptions | Contains a collection of Word.ContentControl objects. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain contents such as images, tables, or paragraphs of formatted text. Currently, only rich text, plain text, checkbox, dropdown list, and combo box content controls are supported. |
| Word.Interfaces.ContentControlCollectionUpdateData | An interface for updating data on the ContentControlCollection object, for use in contentControlCollection.set({ ... }). |
| Word.Interfaces.ContentControlData | An interface describing the data returned by calling contentControl.toJSON(). |
| Word.Interfaces.ContentControlListItemCollectionData | An interface describing the data returned by calling contentControlListItemCollection.toJSON(). |
| Word.Interfaces.ContentControlListItemCollectionLoadOptions | Contains a collection of Word.ContentControlListItem objects that represent the items in a dropdown list or combo box content control. |
| Word.Interfaces.ContentControlListItemCollectionUpdateData | An interface for updating data on the ContentControlListItemCollection object, for use in contentControlListItemCollection.set({ ... }). |
| Word.Interfaces.ContentControlListItemData | An interface describing the data returned by calling contentControlListItem.toJSON(). |
| Word.Interfaces.ContentControlListItemLoadOptions | Represents a list item in a dropdown list or combo box content control. |
| Word.Interfaces.ContentControlListItemUpdateData | An interface for updating data on the ContentControlListItem object, for use in contentControlListItem.set({ ... }). |
| Word.Interfaces.ContentControlLoadOptions | Represents a content control. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain contents such as images, tables, or paragraphs of formatted text. Currently, only rich text, plain text, checkbox, dropdown list, and combo box content controls are supported. |
| Word.Interfaces.ContentControlUpdateData | An interface for updating data on the ContentControl object, for use in contentControl.set({ ... }). |
| Word.Interfaces.CritiqueAnnotationData | An interface describing the data returned by calling critiqueAnnotation.toJSON(). |
| Word.Interfaces.CritiqueAnnotationLoadOptions | Represents an annotation wrapper around critique displayed in the document. |
| Word.Interfaces.CustomPropertyCollectionData | An interface describing the data returned by calling customPropertyCollection.toJSON(). |
| Word.Interfaces.CustomPropertyCollectionLoadOptions | Contains the collection of Word.CustomProperty objects. |
| Word.Interfaces.CustomPropertyCollectionUpdateData | An interface for updating data on the CustomPropertyCollection object, for use in customPropertyCollection.set({ ... }). |
| Word.Interfaces.CustomPropertyData | An interface describing the data returned by calling customProperty.toJSON(). |
| Word.Interfaces.CustomPropertyLoadOptions | Represents a custom property. |
| Word.Interfaces.CustomPropertyUpdateData | An interface for updating data on the CustomProperty object, for use in customProperty.set({ ... }). |
| Word.Interfaces.CustomXmlNodeCollectionData | An interface describing the data returned by calling customXmlNodeCollection.toJSON(). |
| Word.Interfaces.CustomXmlNodeCollectionLoadOptions | Contains a collection of Word.CustomXmlNode objects representing the XML nodes in a document. |
| Word.Interfaces.CustomXmlNodeCollectionUpdateData | An interface for updating data on the CustomXmlNodeCollection object, for use in customXmlNodeCollection.set({ ... }). |
| Word.Interfaces.CustomXmlNodeData | An interface describing the data returned by calling customXmlNode.toJSON(). |
| Word.Interfaces.CustomXmlNodeLoadOptions | Represents an XML node in a tree in the document. The CustomXmlNode object is a member of the Word.CustomXmlNodeCollection object. |
| Word.Interfaces.CustomXmlNodeUpdateData | An interface for updating data on the CustomXmlNode object, for use in customXmlNode.set({ ... }). |
| Word.Interfaces.CustomXmlPartCollectionData | An interface describing the data returned by calling customXmlPartCollection.toJSON(). |
| Word.Interfaces.CustomXmlPartCollectionLoadOptions | Contains the collection of Word.CustomXmlPart objects. |
| Word.Interfaces.CustomXmlPartCollectionUpdateData | An interface for updating data on the CustomXmlPartCollection object, for use in customXmlPartCollection.set({ ... }). |
| Word.Interfaces.CustomXmlPartData | An interface describing the data returned by calling customXmlPart.toJSON(). |
| Word.Interfaces.CustomXmlPartLoadOptions | Represents a custom XML part. |
| Word.Interfaces.CustomXmlPartScopedCollectionData | An interface describing the data returned by calling customXmlPartScopedCollection.toJSON(). |
| Word.Interfaces.CustomXmlPartScopedCollectionLoadOptions | Contains the collection of Word.CustomXmlPart objects with a specific namespace. |
| Word.Interfaces.CustomXmlPartScopedCollectionUpdateData | An interface for updating data on the CustomXmlPartScopedCollection object, for use in customXmlPartScopedCollection.set({ ... }). |
| Word.Interfaces.CustomXmlPartUpdateData | An interface for updating data on the CustomXmlPart object, for use in customXmlPart.set({ ... }). |
| Word.Interfaces.CustomXmlPrefixMappingCollectionData | An interface describing the data returned by calling customXmlPrefixMappingCollection.toJSON(). |
| Word.Interfaces.CustomXmlPrefixMappingCollectionLoadOptions | Represents a collection of Word.CustomXmlPrefixMapping objects. |
| Word.Interfaces.CustomXmlPrefixMappingCollectionUpdateData | An interface for updating data on the CustomXmlPrefixMappingCollection object, for use in customXmlPrefixMappingCollection.set({ ... }). |
| Word.Interfaces.CustomXmlPrefixMappingData | An interface describing the data returned by calling customXmlPrefixMapping.toJSON(). |
| Word.Interfaces.CustomXmlPrefixMappingLoadOptions | Represents a CustomXmlPrefixMapping object. |
| Word.Interfaces.CustomXmlSchemaCollectionData | An interface describing the data returned by calling customXmlSchemaCollection.toJSON(). |
| Word.Interfaces.CustomXmlSchemaCollectionLoadOptions | Represents a collection of Word.CustomXmlSchema objects attached to a data stream. |
| Word.Interfaces.CustomXmlSchemaCollectionUpdateData | An interface for updating data on the CustomXmlSchemaCollection object, for use in customXmlSchemaCollection.set({ ... }). |
| Word.Interfaces.CustomXmlSchemaData | An interface describing the data returned by calling customXmlSchema.toJSON(). |
| Word.Interfaces.CustomXmlSchemaLoadOptions | Represents a schema in a Word.CustomXmlSchemaCollection object. |
| Word.Interfaces.CustomXmlValidationErrorCollectionData | An interface describing the data returned by calling customXmlValidationErrorCollection.toJSON(). |
| Word.Interfaces.CustomXmlValidationErrorCollectionLoadOptions | Represents a collection of Word.CustomXmlValidationError objects. |
| Word.Interfaces.CustomXmlValidationErrorCollectionUpdateData | An interface for updating data on the CustomXmlValidationErrorCollection object, for use in customXmlValidationErrorCollection.set({ ... }). |
| Word.Interfaces.CustomXmlValidationErrorData | An interface describing the data returned by calling customXmlValidationError.toJSON(). |
| Word.Interfaces.CustomXmlValidationErrorLoadOptions | Represents a single validation error in a Word.CustomXmlValidationErrorCollection object. |
| Word.Interfaces.CustomXmlValidationErrorUpdateData | An interface for updating data on the CustomXmlValidationError object, for use in customXmlValidationError.set({ ... }). |
| Word.Interfaces.DatePickerContentControlData | An interface describing the data returned by calling datePickerContentControl.toJSON(). |
| Word.Interfaces.DatePickerContentControlLoadOptions | Represents the DatePickerContentControl object. |
| Word.Interfaces.DatePickerContentControlUpdateData | An interface for updating data on the DatePickerContentControl object, for use in datePickerContentControl.set({ ... }). |
| Word.Interfaces.DocumentCreatedData | An interface describing the data returned by calling documentCreated.toJSON(). |
| Word.Interfaces.DocumentCreatedLoadOptions | The DocumentCreated object is the top level object created by Application.CreateDocument. A DocumentCreated object is a special Document object. |
| Word.Interfaces.DocumentCreatedUpdateData | An interface for updating data on the DocumentCreated object, for use in documentCreated.set({ ... }). |
| Word.Interfaces.DocumentData | An interface describing the data returned by calling document.toJSON(). |
| Word.Interfaces.DocumentLibraryVersionCollectionData | An interface describing the data returned by calling documentLibraryVersionCollection.toJSON(). |
| Word.Interfaces.DocumentLibraryVersionCollectionLoadOptions | Represents the collection of Word.DocumentLibraryVersion objects. |
| Word.Interfaces.DocumentLibraryVersionCollectionUpdateData | An interface for updating data on the DocumentLibraryVersionCollection object, for use in documentLibraryVersionCollection.set({ ... }). |
| Word.Interfaces.DocumentLibraryVersionData | An interface describing the data returned by calling documentLibraryVersion.toJSON(). |
| Word.Interfaces.DocumentLibraryVersionLoadOptions | Represents a document library version. |
| Word.Interfaces.DocumentLoadOptions | The Document object is the top level object. A Document object contains one or more sections, content controls, and the body that contains the contents of the document. |
| Word.Interfaces.DocumentPropertiesData | An interface describing the data returned by calling documentProperties.toJSON(). |
| Word.Interfaces.DocumentPropertiesLoadOptions | Represents document properties. |
| Word.Interfaces.DocumentPropertiesUpdateData | An interface for updating data on the DocumentProperties object, for use in documentProperties.set({ ... }). |
| Word.Interfaces.DocumentUpdateData | An interface for updating data on the Document object, for use in document.set({ ... }). |
| Word.Interfaces.DropCapData | An interface describing the data returned by calling dropCap.toJSON(). |
| Word.Interfaces.DropCapLoadOptions | Represents a dropped capital letter in a Word document. |
| Word.Interfaces.DropDownListContentControlData | An interface describing the data returned by calling dropDownListContentControl.toJSON(). |
| Word.Interfaces.FieldCollectionData | An interface describing the data returned by calling fieldCollection.toJSON(). |
| Word.Interfaces.FieldCollectionLoadOptions | Contains a collection of Word.Field objects. |
| Word.Interfaces.FieldCollectionUpdateData | An interface for updating data on the FieldCollection object, for use in fieldCollection.set({ ... }). |
| Word.Interfaces.FieldData | An interface describing the data returned by calling field.toJSON(). |
| Word.Interfaces.FieldLoadOptions | Represents a field. |
| Word.Interfaces.FieldUpdateData | An interface for updating data on the Field object, for use in field.set({ ... }). |
| Word.Interfaces.FillFormatData | An interface describing the data returned by calling fillFormat.toJSON(). |
| Word.Interfaces.FillFormatLoadOptions | Represents the fill formatting for a shape or text. |
| Word.Interfaces.FillFormatUpdateData | An interface for updating data on the FillFormat object, for use in fillFormat.set({ ... }). |
| Word.Interfaces.FontData | An interface describing the data returned by calling font.toJSON(). |
| Word.Interfaces.FontLoadOptions | Represents a font. |
| Word.Interfaces.FontUpdateData | An interface for updating data on the Font object, for use in font.set({ ... }). |
| Word.Interfaces.FrameCollectionData | An interface describing the data returned by calling frameCollection.toJSON(). |
| Word.Interfaces.FrameCollectionLoadOptions | Represents the collection of Word.Frame objects. |
| Word.Interfaces.FrameCollectionUpdateData | An interface for updating data on the FrameCollection object, for use in frameCollection.set({ ... }). |
| Word.Interfaces.FrameData | An interface describing the data returned by calling frame.toJSON(). |
| Word.Interfaces.FrameLoadOptions | Represents a frame. The Frame object is a member of the Word.FrameCollection object. |
| Word.Interfaces.FrameUpdateData | An interface for updating data on the Frame object, for use in frame.set({ ... }). |
| Word.Interfaces.GlowFormatData | An interface describing the data returned by calling glowFormat.toJSON(). |
| Word.Interfaces.GlowFormatLoadOptions | Represents the glow formatting for the font used by the range of text. |
| Word.Interfaces.GlowFormatUpdateData | An interface for updating data on the GlowFormat object, for use in glowFormat.set({ ... }). |
| Word.Interfaces.GroupContentControlData | An interface describing the data returned by calling groupContentControl.toJSON(). |
| Word.Interfaces.GroupContentControlLoadOptions | Represents the GroupContentControl object. |
| Word.Interfaces.GroupContentControlUpdateData | An interface for updating data on the GroupContentControl object, for use in groupContentControl.set({ ... }). |
| Word.Interfaces.HyperlinkCollectionData | An interface describing the data returned by calling hyperlinkCollection.toJSON(). |
| Word.Interfaces.HyperlinkCollectionLoadOptions | Contains a collection of Word.Hyperlink objects. |
| Word.Interfaces.HyperlinkCollectionUpdateData | An interface for updating data on the HyperlinkCollection object, for use in hyperlinkCollection.set({ ... }). |
| Word.Interfaces.HyperlinkData | An interface describing the data returned by calling hyperlink.toJSON(). |
| Word.Interfaces.HyperlinkLoadOptions | Represents a hyperlink in a Word document. |
| Word.Interfaces.HyperlinkUpdateData | An interface for updating data on the Hyperlink object, for use in hyperlink.set({ ... }). |
| Word.Interfaces.IndexCollectionData | An interface describing the data returned by calling indexCollection.toJSON(). |
| Word.Interfaces.IndexCollectionLoadOptions | A collection of Word.Index objects that represents all the indexes in the document. |
| Word.Interfaces.IndexCollectionUpdateData | An interface for updating data on the IndexCollection object, for use in indexCollection.set({ ... }). |
| Word.Interfaces.IndexData | An interface describing the data returned by calling index.toJSON(). |
| Word.Interfaces.IndexLoadOptions | Represents a single index. The Index object is a member of the Word.IndexCollection. The IndexCollection includes all the indexes in the document. |
| Word.Interfaces.IndexUpdateData | An interface for updating data on the Index object, for use in index.set({ ... }). |
| Word.Interfaces.InlinePictureCollectionData | An interface describing the data returned by calling inlinePictureCollection.toJSON(). |
| Word.Interfaces.InlinePictureCollectionLoadOptions | Contains a collection of Word.InlinePicture objects. |
| Word.Interfaces.InlinePictureCollectionUpdateData | An interface for updating data on the InlinePictureCollection object, for use in inlinePictureCollection.set({ ... }). |
| Word.Interfaces.InlinePictureData | An interface describing the data returned by calling inlinePicture.toJSON(). |
| Word.Interfaces.InlinePictureLoadOptions | Represents an inline picture. |
| Word.Interfaces.InlinePictureUpdateData | An interface for updating data on the InlinePicture object, for use in inlinePicture.set({ ... }). |
| Word.Interfaces.LineFormatData | An interface describing the data returned by calling lineFormat.toJSON(). |
| Word.Interfaces.LineFormatLoadOptions | Represents line and arrowhead formatting. For a line, the LineFormat object contains formatting information for the line itself; for a shape with a border, this object contains formatting information for the shape's border. |
| Word.Interfaces.LineFormatUpdateData | An interface for updating data on the LineFormat object, for use in lineFormat.set({ ... }). |
| Word.Interfaces.LineNumberingData | An interface describing the data returned by calling lineNumbering.toJSON(). |
| Word.Interfaces.LineNumberingLoadOptions | Represents line numbers in the left margin or to the left of each newspaper-style column. |
| Word.Interfaces.LineNumberingUpdateData | An interface for updating data on the LineNumbering object, for use in lineNumbering.set({ ... }). |
| Word.Interfaces.LinkFormatData | An interface describing the data returned by calling linkFormat.toJSON(). |
| Word.Interfaces.LinkFormatLoadOptions | Represents the linking characteristics for an OLE object or picture. |
| Word.Interfaces.LinkFormatUpdateData | An interface for updating data on the LinkFormat object, for use in linkFormat.set({ ... }). |
| Word.Interfaces.ListCollectionData | An interface describing the data returned by calling listCollection.toJSON(). |
| Word.Interfaces.ListCollectionLoadOptions | Contains a collection of Word.List objects. |
| Word.Interfaces.ListCollectionUpdateData | An interface for updating data on the ListCollection object, for use in listCollection.set({ ... }). |
| Word.Interfaces.ListData | An interface describing the data returned by calling list.toJSON(). |
| Word.Interfaces.ListFormatData | An interface describing the data returned by calling listFormat.toJSON(). |
| Word.Interfaces.ListFormatLoadOptions | Represents the list formatting characteristics of a range. |
| Word.Interfaces.ListFormatUpdateData | An interface for updating data on the ListFormat object, for use in listFormat.set({ ... }). |
| Word.Interfaces.ListItemData | An interface describing the data returned by calling listItem.toJSON(). |
| Word.Interfaces.ListItemLoadOptions | Represents the paragraph list item format. |
| Word.Interfaces.ListItemUpdateData | An interface for updating data on the ListItem object, for use in listItem.set({ ... }). |
| Word.Interfaces.ListLevelCollectionData | An interface describing the data returned by calling listLevelCollection.toJSON(). |
| Word.Interfaces.ListLevelCollectionLoadOptions | Contains a collection of Word.ListLevel objects. |
| Word.Interfaces.ListLevelCollectionUpdateData | An interface for updating data on the ListLevelCollection object, for use in listLevelCollection.set({ ... }). |
| Word.Interfaces.ListLevelData | An interface describing the data returned by calling listLevel.toJSON(). |
| Word.Interfaces.ListLevelLoadOptions | Represents a list level. |
| Word.Interfaces.ListLevelUpdateData | An interface for updating data on the ListLevel object, for use in listLevel.set({ ... }). |
| Word.Interfaces.ListLoadOptions | Contains a collection of Word.Paragraph objects. |
| Word.Interfaces.ListTemplateData | An interface describing the data returned by calling listTemplate.toJSON(). |
| Word.Interfaces.ListTemplateLoadOptions | Represents a list template. |
| Word.Interfaces.ListTemplateUpdateData | An interface for updating data on the ListTemplate object, for use in listTemplate.set({ ... }). |
| Word.Interfaces.NoteItemCollectionData | An interface describing the data returned by calling noteItemCollection.toJSON(). |
| Word.Interfaces.NoteItemCollectionLoadOptions | Contains a collection of Word.NoteItem objects. |
| Word.Interfaces.NoteItemCollectionUpdateData | An interface for updating data on the NoteItemCollection object, for use in noteItemCollection.set({ ... }). |
| Word.Interfaces.NoteItemData | An interface describing the data returned by calling noteItem.toJSON(). |
| Word.Interfaces.NoteItemLoadOptions | Represents a footnote or endnote. |
| Word.Interfaces.NoteItemUpdateData | An interface for updating data on the NoteItem object, for use in noteItem.set({ ... }). |
| Word.Interfaces.OleFormatData | An interface describing the data returned by calling oleFormat.toJSON(). |
| Word.Interfaces.OleFormatLoadOptions | Represents the OLE characteristics (other than linking) for an OLE object, ActiveX control, or field. |
| Word.Interfaces.OleFormatUpdateData | An interface for updating data on the OleFormat object, for use in oleFormat.set({ ... }). |
| Word.Interfaces.PageCollectionData | An interface describing the data returned by calling pageCollection.toJSON(). |
| Word.Interfaces.PageCollectionLoadOptions | Represents the collection of page. |
| Word.Interfaces.PageCollectionUpdateData | An interface for updating data on the PageCollection object, for use in pageCollection.set({ ... }). |
| Word.Interfaces.PageData | An interface describing the data returned by calling page.toJSON(). |
| Word.Interfaces.PageLoadOptions | Represents a page in the document. Page objects manage the page layout and content. |
| Word.Interfaces.PageSetupData | An interface describing the data returned by calling pageSetup.toJSON(). |
| Word.Interfaces.PageSetupLoadOptions | Represents the page setup settings for a Word document or section. |
| Word.Interfaces.PageSetupUpdateData | An interface for updating data on the PageSetup object, for use in pageSetup.set({ ... }). |
| Word.Interfaces.PaneCollectionData | An interface describing the data returned by calling paneCollection.toJSON(). |
| Word.Interfaces.PaneCollectionUpdateData | An interface for updating data on the PaneCollection object, for use in paneCollection.set({ ... }). |
| Word.Interfaces.PaneData | An interface describing the data returned by calling pane.toJSON(). |
| Word.Interfaces.ParagraphCollectionData | An interface describing the data returned by calling paragraphCollection.toJSON(). |
| Word.Interfaces.ParagraphCollectionLoadOptions | Contains a collection of Word.Paragraph objects. |
| Word.Interfaces.ParagraphCollectionUpdateData | An interface for updating data on the ParagraphCollection object, for use in paragraphCollection.set({ ... }). |
| Word.Interfaces.ParagraphData | An interface describing the data returned by calling paragraph.toJSON(). |
| Word.Interfaces.ParagraphFormatData | An interface describing the data returned by calling paragraphFormat.toJSON(). |
| Word.Interfaces.ParagraphFormatLoadOptions | Represents a style of paragraph in a document. |
| Word.Interfaces.ParagraphFormatUpdateData | An interface for updating data on the ParagraphFormat object, for use in paragraphFormat.set({ ... }). |
| Word.Interfaces.ParagraphLoadOptions | Represents a single paragraph in a selection, range, content control, or document body. |
| Word.Interfaces.ParagraphUpdateData | An interface for updating data on the Paragraph object, for use in paragraph.set({ ... }). |
| Word.Interfaces.PictureContentControlData | An interface describing the data returned by calling pictureContentControl.toJSON(). |
| Word.Interfaces.PictureContentControlLoadOptions | Represents the PictureContentControl object. |
| Word.Interfaces.PictureContentControlUpdateData | An interface for updating data on the PictureContentControl object, for use in pictureContentControl.set({ ... }). |
| Word.Interfaces.RangeCollectionData | An interface describing the data returned by calling rangeCollection.toJSON(). |
| Word.Interfaces.RangeCollectionLoadOptions | Contains a collection of Word.Range objects. |
| Word.Interfaces.RangeCollectionUpdateData | An interface for updating data on the RangeCollection object, for use in rangeCollection.set({ ... }). |
| Word.Interfaces.RangeData | An interface describing the data returned by calling range.toJSON(). |
| Word.Interfaces.RangeLoadOptions | Represents a contiguous area in a document. |
| Word.Interfaces.RangeUpdateData | An interface for updating data on the Range object, for use in range.set({ ... }). |
| Word.Interfaces.ReflectionFormatData | An interface describing the data returned by calling reflectionFormat.toJSON(). |
| Word.Interfaces.ReflectionFormatLoadOptions | Represents the reflection formatting for a shape in Word. |
| Word.Interfaces.ReflectionFormatUpdateData | An interface for updating data on the ReflectionFormat object, for use in reflectionFormat.set({ ... }). |
| Word.Interfaces.RepeatingSectionContentControlData | An interface describing the data returned by calling repeatingSectionContentControl.toJSON(). |
| Word.Interfaces.RepeatingSectionContentControlLoadOptions | Represents the RepeatingSectionContentControl object. |
| Word.Interfaces.RepeatingSectionContentControlUpdateData | An interface for updating data on the RepeatingSectionContentControl object, for use in repeatingSectionContentControl.set({ ... }). |
| Word.Interfaces.RepeatingSectionItemData | An interface describing the data returned by calling repeatingSectionItem.toJSON(). |
| Word.Interfaces.RepeatingSectionItemLoadOptions | Represents a single item in a Word.RepeatingSectionContentControl. |
| Word.Interfaces.RepeatingSectionItemUpdateData | An interface for updating data on the RepeatingSectionItem object, for use in repeatingSectionItem.set({ ... }). |
| Word.Interfaces.ReviewerCollectionData | An interface describing the data returned by calling reviewerCollection.toJSON(). |
| Word.Interfaces.ReviewerCollectionLoadOptions | A collection of Word.Reviewer objects that represents the reviewers of one or more documents. The ReviewerCollection object contains the names of all reviewers who have reviewed documents opened or edited on a computer. |
| Word.Interfaces.ReviewerCollectionUpdateData | An interface for updating data on the ReviewerCollection object, for use in reviewerCollection.set({ ... }). |
| Word.Interfaces.ReviewerData | An interface describing the data returned by calling reviewer.toJSON(). |
| Word.Interfaces.ReviewerLoadOptions | Represents a single reviewer of a document in which changes have been tracked. The Reviewer object is a member of the Word.ReviewerCollection object. |
| Word.Interfaces.ReviewerUpdateData | An interface for updating data on the Reviewer object, for use in reviewer.set({ ... }). |
| Word.Interfaces.RevisionsFilterData | An interface describing the data returned by calling revisionsFilter.toJSON(). |
| Word.Interfaces.RevisionsFilterLoadOptions | Represents the current settings related to the display of reviewers' comments and revision marks in the document. |
| Word.Interfaces.RevisionsFilterUpdateData | An interface for updating data on the RevisionsFilter object, for use in revisionsFilter.set({ ... }). |
| Word.Interfaces.SearchOptionsData | An interface describing the data returned by calling searchOptions.toJSON(). |
| Word.Interfaces.SearchOptionsLoadOptions | Specifies the options to be included in a search operation. To learn more about how to use search options in the Word JavaScript APIs, read Use search options to find text in your Word add-in. |
| Word.Interfaces.SearchOptionsUpdateData | An interface for updating data on the SearchOptions object, for use in searchOptions.set({ ... }). |
| Word.Interfaces.SectionCollectionData | An interface describing the data returned by calling sectionCollection.toJSON(). |
| Word.Interfaces.SectionCollectionLoadOptions | Contains the collection of the document's Word.Section objects. |
| Word.Interfaces.SectionCollectionUpdateData | An interface for updating data on the SectionCollection object, for use in sectionCollection.set({ ... }). |
| Word.Interfaces.SectionData | An interface describing the data returned by calling section.toJSON(). |
| Word.Interfaces.SectionLoadOptions | Represents a section in a Word document. |
| Word.Interfaces.SectionUpdateData | An interface for updating data on the Section object, for use in section.set({ ... }). |
| Word.Interfaces.SettingCollectionData | An interface describing the data returned by calling settingCollection.toJSON(). |
| Word.Interfaces.SettingCollectionLoadOptions | Contains the collection of Word.Setting objects. |
| Word.Interfaces.SettingCollectionUpdateData | An interface for updating data on the SettingCollection object, for use in settingCollection.set({ ... }). |
| Word.Interfaces.SettingData | An interface describing the data returned by calling setting.toJSON(). |
| Word.Interfaces.SettingLoadOptions | Represents a setting of the add-in. |
| Word.Interfaces.SettingUpdateData | An interface for updating data on the Setting object, for use in setting.set({ ... }). |
| Word.Interfaces.ShadingData | An interface describing the data returned by calling shading.toJSON(). |
| Word.Interfaces.ShadingLoadOptions | Represents the shading object. |
| Word.Interfaces.ShadingUniversalData | An interface describing the data returned by calling shadingUniversal.toJSON(). |
| Word.Interfaces.ShadingUniversalLoadOptions | Represents the ShadingUniversal object, which manages shading for a range, paragraph, frame, or table. |
| Word.Interfaces.ShadingUniversalUpdateData | An interface for updating data on the ShadingUniversal object, for use in shadingUniversal.set({ ... }). |
| Word.Interfaces.ShadingUpdateData | An interface for updating data on the Shading object, for use in shading.set({ ... }). |
| Word.Interfaces.ShadowFormatData | An interface describing the data returned by calling shadowFormat.toJSON(). |
| Word.Interfaces.ShadowFormatLoadOptions | Represents the shadow formatting for a shape or text in Word. |
| Word.Interfaces.ShadowFormatUpdateData | An interface for updating data on the ShadowFormat object, for use in shadowFormat.set({ ... }). |
| Word.Interfaces.ShapeCollectionData | An interface describing the data returned by calling shapeCollection.toJSON(). |
| Word.Interfaces.ShapeCollectionLoadOptions | Contains a collection of Word.Shape objects. Currently, only the following shapes are supported: text boxes, geometric shapes, groups, pictures, and canvases. |
| Word.Interfaces.ShapeCollectionUpdateData | An interface for updating data on the ShapeCollection object, for use in shapeCollection.set({ ... }). |
| Word.Interfaces.ShapeData | An interface describing the data returned by calling shape.toJSON(). |
| Word.Interfaces.ShapeFillData | An interface describing the data returned by calling shapeFill.toJSON(). |
| Word.Interfaces.ShapeFillLoadOptions | Represents the fill formatting of a shape object. |
| Word.Interfaces.ShapeFillUpdateData | An interface for updating data on the ShapeFill object, for use in shapeFill.set({ ... }). |
| Word.Interfaces.ShapeGroupData | An interface describing the data returned by calling shapeGroup.toJSON(). |
| Word.Interfaces.ShapeGroupLoadOptions | Represents a shape group in the document. To get the corresponding Shape object, use ShapeGroup.shape. |
| Word.Interfaces.ShapeGroupUpdateData | An interface for updating data on the ShapeGroup object, for use in shapeGroup.set({ ... }). |
| Word.Interfaces.ShapeLoadOptions | Represents a shape in the header, footer, or document body. Currently, only the following shapes are supported: text boxes, geometric shapes, groups, pictures, and canvases. |
| Word.Interfaces.ShapeTextWrapData | An interface describing the data returned by calling shapeTextWrap.toJSON(). |
| Word.Interfaces.ShapeTextWrapLoadOptions | Represents all the properties for wrapping text around a shape. |
| Word.Interfaces.ShapeTextWrapUpdateData | An interface for updating data on the ShapeTextWrap object, for use in shapeTextWrap.set({ ... }). |
| Word.Interfaces.ShapeUpdateData | An interface for updating data on the Shape object, for use in shape.set({ ... }). |
| Word.Interfaces.SourceCollectionData | An interface describing the data returned by calling sourceCollection.toJSON(). |
| Word.Interfaces.SourceCollectionLoadOptions | Represents a collection of Word.Source objects. |
| Word.Interfaces.SourceCollectionUpdateData | An interface for updating data on the SourceCollection object, for use in sourceCollection.set({ ... }). |
| Word.Interfaces.SourceData | An interface describing the data returned by calling source.toJSON(). |
| Word.Interfaces.SourceLoadOptions | Represents an individual source, such as a book, journal article, or interview. |
| Word.Interfaces.StyleCollectionData | An interface describing the data returned by calling styleCollection.toJSON(). |
| Word.Interfaces.StyleCollectionLoadOptions | Contains a collection of Word.Style objects. |
| Word.Interfaces.StyleCollectionUpdateData | An interface for updating data on the StyleCollection object, for use in styleCollection.set({ ... }). |
| Word.Interfaces.StyleData | An interface describing the data returned by calling style.toJSON(). |
| Word.Interfaces.StyleLoadOptions | Represents a style in a Word document. |
| Word.Interfaces.StyleUpdateData | An interface for updating data on the Style object, for use in style.set({ ... }). |
| Word.Interfaces.TableBorderData | An interface describing the data returned by calling tableBorder.toJSON(). |
| Word.Interfaces.TableBorderLoadOptions | Specifies the border style. |
| Word.Interfaces.TableBorderUpdateData | An interface for updating data on the TableBorder object, for use in tableBorder.set({ ... }). |
| Word.Interfaces.TableCellCollectionData | An interface describing the data returned by calling tableCellCollection.toJSON(). |
| Word.Interfaces.TableCellCollectionLoadOptions | Contains the collection of the document's TableCell objects. |
| Word.Interfaces.TableCellCollectionUpdateData | An interface for updating data on the TableCellCollection object, for use in tableCellCollection.set({ ... }). |
| Word.Interfaces.TableCellData | An interface describing the data returned by calling tableCell.toJSON(). |
| Word.Interfaces.TableCellLoadOptions | Represents a table cell in a Word document. |
| Word.Interfaces.TableCellUpdateData | An interface for updating data on the TableCell object, for use in tableCell.set({ ... }). |
| Word.Interfaces.TableCollectionData | An interface describing the data returned by calling tableCollection.toJSON(). |
| Word.Interfaces.TableCollectionLoadOptions | Contains the collection of the document's Table objects. |
| Word.Interfaces.TableCollectionUpdateData | An interface for updating data on the TableCollection object, for use in tableCollection.set({ ... }). |
| Word.Interfaces.TableColumnCollectionData | An interface describing the data returned by calling tableColumnCollection.toJSON(). |
| Word.Interfaces.TableColumnCollectionLoadOptions | Represents a collection of Word.TableColumn objects in a Word document. |
| Word.Interfaces.TableColumnCollectionUpdateData | An interface for updating data on the TableColumnCollection object, for use in tableColumnCollection.set({ ... }). |
| Word.Interfaces.TableColumnData | An interface describing the data returned by calling tableColumn.toJSON(). |
| Word.Interfaces.TableColumnLoadOptions | Represents a table column in a Word document. |
| Word.Interfaces.TableColumnUpdateData | An interface for updating data on the TableColumn object, for use in tableColumn.set({ ... }). |
| Word.Interfaces.TableData | An interface describing the data returned by calling table.toJSON(). |
| Word.Interfaces.TableLoadOptions | Represents a table in a Word document. |
| Word.Interfaces.TableRowCollectionData | An interface describing the data returned by calling tableRowCollection.toJSON(). |
| Word.Interfaces.TableRowCollectionLoadOptions | Contains the collection of the document's TableRow objects. |
| Word.Interfaces.TableRowCollectionUpdateData | An interface for updating data on the TableRowCollection object, for use in tableRowCollection.set({ ... }). |
| Word.Interfaces.TableRowData | An interface describing the data returned by calling tableRow.toJSON(). |
| Word.Interfaces.TableRowLoadOptions | Represents a row in a Word document. |
| Word.Interfaces.TableRowUpdateData | An interface for updating data on the TableRow object, for use in tableRow.set({ ... }). |
| Word.Interfaces.TableStyleData | An interface describing the data returned by calling tableStyle.toJSON(). |
| Word.Interfaces.TableStyleLoadOptions | Represents the TableStyle object. |
| Word.Interfaces.TableStyleUpdateData | An interface for updating data on the TableStyle object, for use in tableStyle.set({ ... }). |
| Word.Interfaces.TableUpdateData | An interface for updating data on the Table object, for use in table.set({ ... }). |
| Word.Interfaces.TabStopCollectionData | An interface describing the data returned by calling tabStopCollection.toJSON(). |
| Word.Interfaces.TabStopCollectionLoadOptions | Represents a collection of tab stops in a Word document. |
| Word.Interfaces.TabStopCollectionUpdateData | An interface for updating data on the TabStopCollection object, for use in tabStopCollection.set({ ... }). |
| Word.Interfaces.TabStopData | An interface describing the data returned by calling tabStop.toJSON(). |
| Word.Interfaces.TabStopLoadOptions | Represents a tab stop in a Word document. |
| Word.Interfaces.TemplateCollectionData | An interface describing the data returned by calling templateCollection.toJSON(). |
| Word.Interfaces.TemplateCollectionLoadOptions | Contains a collection of Word.Template objects that represent all the templates that are currently available. This collection includes open templates, templates attached to open documents, and global templates loaded in the Templates and Add-ins dialog box. To learn how to access this dialog in the Word UI, see Load or unload a template or add-in program. |
| Word.Interfaces.TemplateCollectionUpdateData | An interface for updating data on the TemplateCollection object, for use in templateCollection.set({ ... }). |
| Word.Interfaces.TemplateData | An interface describing the data returned by calling template.toJSON(). |
| Word.Interfaces.TemplateLoadOptions | Represents a document template. |
| Word.Interfaces.TemplateUpdateData | An interface for updating data on the Template object, for use in template.set({ ... }). |
| Word.Interfaces.TextColumnCollectionData | An interface describing the data returned by calling textColumnCollection.toJSON(). |
| Word.Interfaces.TextColumnCollectionLoadOptions | A collection of Word.TextColumn objects that represent all the columns of text in the document or a section of the document. |
| Word.Interfaces.TextColumnCollectionUpdateData | An interface for updating data on the TextColumnCollection object, for use in textColumnCollection.set({ ... }). |
| Word.Interfaces.TextColumnData | An interface describing the data returned by calling textColumn.toJSON(). |
| Word.Interfaces.TextColumnLoadOptions | Represents a single text column in a section. |
| Word.Interfaces.TextColumnUpdateData | An interface for updating data on the TextColumn object, for use in textColumn.set({ ... }). |
| Word.Interfaces.TextFrameData | An interface describing the data returned by calling textFrame.toJSON(). |
| Word.Interfaces.TextFrameLoadOptions | Represents the text frame of a shape object. |
| Word.Interfaces.TextFrameUpdateData | An interface for updating data on the TextFrame object, for use in textFrame.set({ ... }). |
| Word.Interfaces.ThreeDimensionalFormatData | An interface describing the data returned by calling threeDimensionalFormat.toJSON(). |
| Word.Interfaces.ThreeDimensionalFormatLoadOptions | Represents a shape's three-dimensional formatting. |
| Word.Interfaces.ThreeDimensionalFormatUpdateData | An interface for updating data on the ThreeDimensionalFormat object, for use in threeDimensionalFormat.set({ ... }). |
| Word.Interfaces.TrackedChangeCollectionData | An interface describing the data returned by calling trackedChangeCollection.toJSON(). |
| Word.Interfaces.TrackedChangeCollectionLoadOptions | Contains a collection of Word.TrackedChange objects. |
| Word.Interfaces.TrackedChangeCollectionUpdateData | An interface for updating data on the TrackedChangeCollection object, for use in trackedChangeCollection.set({ ... }). |
| Word.Interfaces.TrackedChangeData | An interface describing the data returned by calling trackedChange.toJSON(). |
| Word.Interfaces.TrackedChangeLoadOptions | Represents a tracked change in a Word document. |
| Word.Interfaces.ViewData | An interface describing the data returned by calling view.toJSON(). |
| Word.Interfaces.ViewLoadOptions | Contains the view attributes (such as show all, field shading, and table gridlines) for a window or pane. |
| Word.Interfaces.ViewUpdateData | An interface for updating data on the View object, for use in view.set({ ... }). |
| Word.Interfaces.WindowCollectionData | An interface describing the data returned by calling windowCollection.toJSON(). |
| Word.Interfaces.WindowCollectionLoadOptions | Represents the collection of window objects. |
| Word.Interfaces.WindowCollectionUpdateData | An interface for updating data on the WindowCollection object, for use in windowCollection.set({ ... }). |
| Word.Interfaces.WindowData | An interface describing the data returned by calling window.toJSON(). |
| Word.Interfaces.WindowLoadOptions | Represents the window that displays the document. A window can be split to contain multiple reading panes. |
| Word.Interfaces.WindowUpdateData | An interface for updating data on the Window object, for use in window.set({ ... }). |
| Word.Interfaces.XmlMappingData | An interface describing the data returned by calling xmlMapping.toJSON(). |
| Word.Interfaces.XmlMappingLoadOptions | Represents the XML mapping on a Word.ContentControl object between custom XML and that content control. An XML mapping is a link between the text in a content control and an XML element in the custom XML data store for this document. |
| Word.Interfaces.XmlMappingUpdateData | An interface for updating data on the XmlMapping object, for use in xmlMapping.set({ ... }). |
| Word.ListFormatCountNumberedItemsOptions | Represents options for counting numbered items in a range. |
| Word.ListTemplateApplyOptions | Represents options for applying a list template to a range. |
| Word.ParagraphAddedEventArgs | Provides information about the paragraphs that raised the paragraphAdded event. |
| Word.ParagraphChangedEventArgs | Provides information about the paragraphs that raised the paragraphChanged event. |
| Word.ParagraphDeletedEventArgs | Provides information about the paragraphs that raised the paragraphDeleted event. |
| Word.TabStopAddOptions | Specifies the options for adding to a Word.TabStopCollection object. |
| Word.TextColumnAddOptions | Represents options for a new text column in a document or section of a document. |
| Word.WindowCloseOptions | The options that define whether to save changes before closing and whether to route the document. |
| Word.WindowPageScrollOptions | The options for scrolling through the specified pane or window page by page. |
| Word.WindowScrollOptions | The options that scrolls a window or pane by the specified number of units defined by the calling method. |
| Word.XmlSetMappingOptions | The options that define the prefix mapping and the source of the custom XML data. |
