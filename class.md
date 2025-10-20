
## Classes

| Class | Description |
|-------|-------------|
| Word.Annotation | Represents an annotation attached to a paragraph. |
| Word.AnnotationCollection | Contains a collection of Word.Annotation objects. |
| Word.Application | Represents the application object. |
| Word.Bibliography | Represents the list of available sources attached to the document (in the current list) or the list of sources available in the application (in the master list). |
| Word.Body | Represents the body of a document or a section. |
| Word.Bookmark | Represents a single bookmark in a document, selection, or range. The Bookmark object is a member of the Bookmark collection. The Word.BookmarkCollection includes all the bookmarks listed in the Bookmark dialog box (Insert menu). |
| Word.BookmarkCollection | A collection of Word.Bookmark objects that represent the bookmarks in the specified selection, range, or document. |
| Word.Border | Represents the Border object for text, a paragraph, or a table. |
| Word.BorderCollection | Represents the collection of border styles. |
| Word.BorderUniversal | Represents the BorderUniversal object, which manages borders for a range, paragraph, table, or frame. |
| Word.BorderUniversalCollection | Represents the collection of Word.BorderUniversal objects. |
| Word.Break | Represents a break in a Word document. This could be a page, column, or section break. |
| Word.BreakCollection | Contains a collection of Word.Break objects. |
| Word.BuildingBlock | Represents a building block in a template. A building block is pre-built content, similar to autotext, that may contain text, images, and formatting. |
| Word.BuildingBlockCategory | Represents a category of building blocks in a Word document. |
| Word.BuildingBlockCategoryCollection | Represents a collection of Word.BuildingBlockCategory objects in a Word document. |
| Word.BuildingBlockCollection | Represents a collection of Word.BuildingBlock objects for a specific building block type and category in a template. |
| Word.BuildingBlockEntryCollection | Represents a collection of building block entries in a Word template. |
| Word.BuildingBlockGalleryContentControl | Represents the BuildingBlockGalleryContentControl object. |
| Word.BuildingBlockTypeItem | Represents a type of building block in a Word document. |
| Word.BuildingBlockTypeItemCollection | Represents a collection of building block types in a Word document. |
| Word.Canvas | Represents a canvas in the document. To get the corresponding Shape object, use Canvas.shape. |
| Word.CheckboxContentControl | The data specific to content controls of type CheckBox. |
| Word.ColorFormat | Represents the color formatting of a shape or text in Word. |
| Word.ComboBoxContentControl | The data specific to content controls of type 'ComboBox'. |
| Word.Comment | Represents a comment in the document. |
| Word.CommentCollection | Contains a collection of Word.Comment objects. |
| Word.CommentContentRange | |
| Word.CommentReply | Represents a comment reply in the document. |
| Word.CommentReplyCollection | Contains a collection of Word.CommentReply objects. Represents all comment replies in one comment thread. |
| Word.ContentControl | Represents a content control. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain contents such as images, tables, or paragraphs of formatted text. Currently, only rich text, plain text, checkbox, dropdown list, and combo box content controls are supported. |
| Word.ContentControlCollection | Contains a collection of Word.ContentControl objects. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain contents such as images, tables, or paragraphs of formatted text. Currently, only rich text, plain text, checkbox, dropdown list, and combo box content controls are supported. |
| Word.ContentControlListItem | Represents a list item in a dropdown list or combo box content control. |
| Word.ContentControlListItemCollection | Contains a collection of Word.ContentControlListItem objects that represent the items in a dropdown list or combo box content control. |
| Word.CritiqueAnnotation | Represents an annotation wrapper around critique displayed in the document. |
| Word.CustomProperty | Represents a custom property. |
| Word.CustomPropertyCollection | Contains the collection of Word.CustomProperty objects. |
| Word.CustomXmlNode | Represents an XML node in a tree in the document. The CustomXmlNode object is a member of the Word.CustomXmlNodeCollection object. |
| Word.CustomXmlNodeCollection | Contains a collection of Word.CustomXmlNode objects representing the XML nodes in a document. |
| Word.CustomXmlPart | Represents a custom XML part. |
| Word.CustomXmlPartCollection | Contains the collection of Word.CustomXmlPart objects. |
| Word.CustomXmlPartScopedCollection | Contains the collection of Word.CustomXmlPart objects with a specific namespace. |
| Word.CustomXmlPrefixMapping | Represents a CustomXmlPrefixMapping object. |
| Word.CustomXmlPrefixMappingCollection | Represents a collection of Word.CustomXmlPrefixMapping objects. |
| Word.CustomXmlSchema | Represents a schema in a Word.CustomXmlSchemaCollection object. |
| Word.CustomXmlSchemaCollection | Represents a collection of Word.CustomXmlSchema objects attached to a data stream. |
| Word.CustomXmlValidationError | Represents a single validation error in a Word.CustomXmlValidationErrorCollection object. |
| Word.CustomXmlValidationErrorCollection | Represents a collection of Word.CustomXmlValidationError objects. |
| Word.DatePickerContentControl | Represents the DatePickerContentControl object. |
| Word.Document | The Document object is the top level object. A Document object contains one or more sections, content controls, and the body that contains the contents of the document. |
| Word.DocumentCreated | The DocumentCreated object is the top level object created by Application.CreateDocument. A DocumentCreated object is a special Document object. |
| Word.DocumentLibraryVersion | Represents a document library version. |
| Word.DocumentLibraryVersionCollection | Represents the collection of Word.DocumentLibraryVersion objects. |
| Word.DocumentProperties | Represents document properties. |
| Word.DropCap | Represents a dropped capital letter in a Word document. |
| Word.DropDownListContentControl | The data specific to content controls of type DropDownList. |
| Word.Field | Represents a field. |
| Word.FieldCollection | Contains a collection of Word.Field objects. |
| Word.FillFormat | Represents the fill formatting for a shape or text. |
| Word.Font | Represents a font. |
| Word.Frame | Represents a frame. The Frame object is a member of the Word.FrameCollection object. |
| Word.FrameCollection | Represents the collection of Word.Frame objects. |
| Word.GlowFormat | Represents the glow formatting for the font used by the range of text. |
| Word.GroupContentControl | Represents the GroupContentControl object. |
| Word.Hyperlink | Represents a hyperlink in a Word document. |
| Word.HyperlinkCollection | Contains a collection of Word.Hyperlink objects. |
| Word.Index | Represents a single index. The Index object is a member of the Word.IndexCollection. The IndexCollection includes all the indexes in the document. |
| Word.IndexCollection | A collection of Word.Index objects that represents all the indexes in the document. |
| Word.InlinePicture | Represents an inline picture. |
| Word.InlinePictureCollection | Contains a collection of Word.InlinePicture objects. |
| Word.LineFormat | Represents line and arrowhead formatting. For a line, the LineFormat object contains formatting information for the line itself; for a shape with a border, this object contains formatting information for the shape's border. |
| Word.LineNumbering | Represents line numbers in the left margin or to the left of each newspaper-style column. |
| Word.LinkFormat | Represents the linking characteristics for an OLE object or picture. |
| Word.List | Contains a collection of Word.Paragraph objects. |
| Word.ListCollection | Contains a collection of Word.List objects. |
| Word.ListFormat | Represents the list formatting characteristics of a range. |
| Word.ListItem | Represents the paragraph list item format. |
| Word.ListLevel | Represents a list level. |
| Word.ListLevelCollection | Contains a collection of Word.ListLevel objects. |
| Word.ListTemplate | Represents a list template. |
| Word.NoteItem | Represents a footnote or endnote. |
| Word.NoteItemCollection | Contains a collection of Word.NoteItem objects. |
| Word.OleFormat | Represents the OLE characteristics (other than linking) for an OLE object, ActiveX control, or field. |
| Word.Page | Represents a page in the document. Page objects manage the page layout and content. |
| Word.PageCollection | Represents the collection of page. |
| Word.PageSetup | Represents the page setup settings for a Word document or section. |
| Word.Pane | Represents a window pane. The Pane object is a member of the pane collection. The pane collection includes all the window panes for a single window. |
| Word.PaneCollection | Represents the collection of pane. |
| Word.Paragraph | Represents a single paragraph in a selection, range, content control, or document body. |
| Word.ParagraphCollection | Contains a collection of Word.Paragraph objects. |
| Word.ParagraphFormat | Represents a style of paragraph in a document. |
| Word.PictureContentControl | Represents the PictureContentControl object. |
| Word.Range | Represents a contiguous area in a document. |
| Word.RangeCollection | Contains a collection of Word.Range objects. |
| Word.ReflectionFormat | Represents the reflection formatting for a shape in Word. |
| Word.RepeatingSectionContentControl | Represents the RepeatingSectionContentControl object. |
| Word.RepeatingSectionItem | Represents a single item in a Word.RepeatingSectionContentControl. |
| Word.RepeatingSectionItemCollection | Represents a collection of Word.RepeatingSectionItem objects in a Word document. |
| Word.RequestContext | The RequestContext object facilitates requests to the Word application. Since the Office add-in and the Word application run in two different processes, the request context is required to get access to the Word object model from the add-in. |
| Word.Reviewer | Represents a single reviewer of a document in which changes have been tracked. The Reviewer object is a member of the Word.ReviewerCollection object. |
| Word.ReviewerCollection | A collection of Word.Reviewer objects that represents the reviewers of one or more documents. The ReviewerCollection object contains the names of all reviewers who have reviewed documents opened or edited on a computer. |
| Word.RevisionsFilter | Represents the current settings related to the display of reviewers' comments and revision marks in the document. |
| Word.SearchOptions | Specifies the options to be included in a search operation. To learn more about how to use search options in the Word JavaScript APIs, read Use search options to find text in your Word add-in. |
| Word.Section | Represents a section in a Word document. |
| Word.SectionCollection | Contains the collection of the document's Word.Section objects. |
| Word.Setting | Represents a setting of the add-in. |
| Word.SettingCollection | Contains the collection of Word.Setting objects. |
| Word.Shading | Represents the shading object. |
| Word.ShadingUniversal | Represents the ShadingUniversal object, which manages shading for a range, paragraph, frame, or table. |
| Word.ShadowFormat | Represents the shadow formatting for a shape or text in Word. |
| Word.Shape | Represents a shape in the header, footer, or document body. Currently, only the following shapes are supported: text boxes, geometric shapes, groups, pictures, and canvases. |
| Word.ShapeCollection | Contains a collection of Word.Shape objects. Currently, only the following shapes are supported: text boxes, geometric shapes, groups, pictures, and canvases. |
| Word.ShapeFill | Represents the fill formatting of a shape object. |
| Word.ShapeGroup | Represents a shape group in the document. To get the corresponding Shape object, use ShapeGroup.shape. |
| Word.ShapeTextWrap | Represents all the properties for wrapping text around a shape. |
| Word.Source | Represents an individual source, such as a book, journal article, or interview. |
| Word.SourceCollection | Represents a collection of Word.Source objects. |
| Word.Style | Represents a style in a Word document. |
| Word.StyleCollection | Contains a collection of Word.Style objects. |
| Word.Table | Represents a table in a Word document. |
| Word.TableBorder | Specifies the border style. |
| Word.TableCell | Represents a table cell in a Word document. |
| Word.TableCellCollection | Contains the collection of the document's TableCell objects. |
| Word.TableCollection | Contains the collection of the document's Table objects. |
| Word.TableColumn | Represents a table column in a Word document. |
| Word.TableColumnCollection | Represents a collection of Word.TableColumn objects in a Word document. |
| Word.TableRow | Represents a row in a Word document. |
| Word.TableRowCollection | Contains the collection of the document's TableRow objects. |
| Word.TableStyle | Represents the TableStyle object. |
| Word.TabStop | Represents a tab stop in a Word document. |
| Word.TabStopCollection | Represents a collection of tab stops in a Word document. |
| Word.Template | Represents a document template. |
| Word.TemplateCollection | Contains a collection of Word.Template objects that represent all the templates that are currently available. This collection includes open templates, templates attached to open documents, and global templates loaded in the Templates and Add-ins dialog box. To learn how to access this dialog in the Word UI, see Load or unload a template or add-in program. |
| Word.TextColumn | Represents a single text column in a section. |
| Word.TextColumnCollection | A collection of Word.TextColumn objects that represent all the columns of text in the document or a section of the document. |
| Word.TextFrame | Represents the text frame of a shape object. |
| Word.ThreeDimensionalFormat | Represents a shape's three-dimensional formatting. |
| Word.TrackedChange | Represents a tracked change in a Word document. |
| Word.TrackedChangeCollection | Contains a collection of Word.TrackedChange objects. |
| Word.View | Contains the view attributes (such as show all, field shading, and table gridlines) for a window or pane. |
| Word.Window | Represents the window that displays the document. A window can be split to contain multiple reading panes. |
| Word.WindowCollection | Represents the collection of window objects. |
| Word.XmlMapping | Represents the XML mapping on a Word.ContentControl object between custom XML and that content control. An XML mapping is a link between the text in a content control and an XML element in the custom XML data store for this document. |
