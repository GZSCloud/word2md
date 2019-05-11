## Document objects
The main Document and related objects.

Document constructor
docx.Document(docx=None)[source]
Return a Document object loaded from docx, where docx can be either a path to a .docx file (a string) or a file-like object. If docx is missing or None, the built-in default document “template” is loaded.

Document objects
class docx.document.Document[source]
WordprocessingML (WML) document.

Not intended to be constructed directly. Use docx.Document() to open or create a document.

add_heading(text=u'', level=1)[source]
Return a heading paragraph newly added to the end of the document.

The heading paragraph will contain text and have its paragraph style determined by level. If level is 0, the style is set to Title. If level is 1 (or omitted), Heading 1 is used. Otherwise the style is set to Heading {level}. Raises ValueError if level is outside the range 0-9.

add_page_break()[source]
Return newly Paragraph object containing only a page break.

add_paragraph(text=u'', style=None)[source]
Return a paragraph newly added to the end of the document, populated with text and having paragraph style style. text can contain tab (\t) characters, which are converted to the appropriate XML form for a tab. text can also include newline (\n) or carriage return (\r) characters, each of which is converted to a line break.

add_picture(image_path_or_stream, width=None, height=None)[source]
Return a new picture shape added in its own paragraph at the end of the document. The picture contains the image at image_path_or_stream, scaled based on width and height. If neither width nor height is specified, the picture appears at its native size. If only one is specified, it is used to compute a scaling factor that is then applied to the unspecified dimension, preserving the aspect ratio of the image. The native size of the picture is calculated using the dots-per-inch (dpi) value specified in the image file, defaulting to 72 dpi if no value is specified, as is often the case.

add_section(start_type=2)[source]
Return a Section object representing a new section added at the end of the document. The optional start_type argument must be a member of the WD_SECTION_START enumeration, and defaults to WD_SECTION.NEW_PAGE if not provided.

add_table(rows, cols, style=None)[source]
Add a table having row and column counts of rows and cols respectively and table style of style. style may be a paragraph style object or a paragraph style name. If style is None, the table inherits the default table style of the document.

core_properties
A CoreProperties object providing read/write access to the core properties of this document.

inline_shapes
An InlineShapes object providing access to the inline shapes in this document. An inline shape is a graphical object, such as a picture, contained in a run of text and behaving like a character glyph, being flowed like other text in a paragraph.

paragraphs
A list of Paragraph instances corresponding to the paragraphs in the document, in document order. Note that paragraphs within revision marks such as <w:ins> or <w:del> do not appear in this list.

part
The DocumentPart object of this document.

save(path_or_stream)[source]
Save this document to path_or_stream, which can be either a path to a filesystem location (a string) or a file-like object.

sections
Sections object providing access to each section in this document.

settings
A Settings object providing access to the document-level settings for this document.

styles
A Styles object providing access to the styles in this document.

tables
A list of Table instances corresponding to the tables in the document, in document order. Note that only tables appearing at the top level of the document appear in this list; a table nested inside a table cell does not appear. A table within revision marks such as <w:ins> or <w:del> will also not appear in the list.

CoreProperties objects
Each Document object provides access to its CoreProperties object via its core_properties attribute. A CoreProperties object provides read/write access to the so-called core properties for the document. The core properties are author, category, comments, content_status, created, identifier, keywords, language, last_modified_by, last_printed, modified, revision, subject, title, and version.

Each property is one of three types, str, datetime.datetime, or int. String properties are limited in length to 255 characters and return an empty string (‘’) if not set. Date properties are assigned and returned as datetime.datetime objects without timezone, i.e. in UTC. Any timezone conversions are the responsibility of the client. Date properties return None if not set.

python-docx does not automatically set any of the document core properties other than to add a core properties part to a presentation that doesn’t have one (very uncommon). If python-docx adds a core properties part, it contains default values for the title, last_modified_by, revision, and modified properties. Client code should update properties like revision and last_modified_by if that behavior is desired.

class docx.opc.coreprops.CoreProperties[source]
author
string – An entity primarily responsible for making the content of the resource.

category
string – A categorization of the content of this package. Example values might include: Resume, Letter, Financial Forecast, Proposal, or Technical Presentation.

comments
string – An account of the content of the resource.

content_status
string – completion status of the document, e.g. ‘draft’

created
datetime – time of intial creation of the document

identifier
string – An unambiguous reference to the resource within a given context, e.g. ISBN.

keywords
string – descriptive words or short phrases likely to be used as search terms for this document

language
string – language the document is written in

last_modified_by
string – name or other identifier (such as email address) of person who last modified the document

last_printed
datetime – time the document was last printed

modified
datetime – time the document was last modified

revision
int – number of this revision, incremented by Word each time the document is saved. Note however python-docx does not automatically increment the revision number when it saves a document.

subject
string – The topic of the content of the resource.

title
string – The name given to the resource.

version
string – free-form version string

Style-related objects
A style is used to collect a set of formatting properties under a single name and apply those properties to a content object all at once. This promotes formatting consistency thoroughout a document and across related documents and allows formatting changes to be made globally by changing the definition in the appropriate style.

Styles objects
class docx.styles.styles.Styles[source]
Provides access to the styles defined in a document.

Accessed using the Document.styles property. Supports len(), iteration, and dictionary-style access by style name.

add_style(name, style_type, builtin=False)[source]
Return a newly added style object of style_type and identified by name. A builtin style can be defined by passing True for the optional builtin argument.

default(style_type)[source]
Return the default style for style_type or None if no default is defined for that type (not common).

element
The lxml element proxied by this object.

latent_styles
A LatentStyles object providing access to the default behaviors for latent styles and the collection of _LatentStyle objects that define overrides of those defaults for a particular named latent style.

BaseStyle objects
class docx.styles.style.BaseStyle[source]
Base class for the various types of style object, paragraph, character, table, and numbering. These properties and methods are inherited by all style objects.

builtin
Read-only. True if this style is a built-in style. False indicates it is a custom (user-defined) style. Note this value is based on the presence of a customStyle attribute in the XML, not on specific knowledge of which styles are built into Word.

delete()[source]
Remove this style definition from the document. Note that calling this method does not remove or change the style applied to any document content. Content items having the deleted style will be rendered using the default style, as is any content with a style not defined in the document.

element
The lxml element proxied by this object.

hidden
True if display of this style in the style gallery and list of recommended styles is suppressed. False otherwise. In order to be shown in the style gallery, this value must be False and quick_style must be True.

locked
Read/write Boolean. True if this style is locked. A locked style does not appear in the styles panel or the style gallery and cannot be applied to document content. This behavior is only active when formatting protection is turned on for the document (via the Developer menu).

name
The UI name of this style.

priority
The integer sort key governing display sequence of this style in the Word UI. None indicates no setting is defined, causing Word to use the default value of 0. Style name is used as a secondary sort key to resolve ordering of styles having the same priority value.

quick_style
True if this style should be displayed in the style gallery when hidden is False. Read/write Boolean.

type
Member of WD_STYLE_TYPE corresponding to the type of this style, e.g. WD_STYLE_TYPE.PARAGRAPH.

unhide_when_used
True if an application should make this style visible the next time it is applied to content. False otherwise. Note that python-docx does not automatically unhide a style having True for this attribute when it is applied to content.

_CharacterStyle objects
class docx.styles.style._CharacterStyle[source]
Bases: docx.styles.style.BaseStyle

A character style. A character style is applied to a Run object and primarily provides character-level formatting via the Font object in its font property.

base_style
Style object this style inherits from or None if this style is not based on another style.

builtin
Read-only. True if this style is a built-in style. False indicates it is a custom (user-defined) style. Note this value is based on the presence of a customStyle attribute in the XML, not on specific knowledge of which styles are built into Word.

delete()
Remove this style definition from the document. Note that calling this method does not remove or change the style applied to any document content. Content items having the deleted style will be rendered using the default style, as is any content with a style not defined in the document.

font
The Font object providing access to the character formatting properties for this style, such as font name and size.

hidden
True if display of this style in the style gallery and list of recommended styles is suppressed. False otherwise. In order to be shown in the style gallery, this value must be False and quick_style must be True.

locked
Read/write Boolean. True if this style is locked. A locked style does not appear in the styles panel or the style gallery and cannot be applied to document content. This behavior is only active when formatting protection is turned on for the document (via the Developer menu).

name
The UI name of this style.

priority
The integer sort key governing display sequence of this style in the Word UI. None indicates no setting is defined, causing Word to use the default value of 0. Style name is used as a secondary sort key to resolve ordering of styles having the same priority value.

quick_style
True if this style should be displayed in the style gallery when hidden is False. Read/write Boolean.

unhide_when_used
True if an application should make this style visible the next time it is applied to content. False otherwise. Note that python-docx does not automatically unhide a style having True for this attribute when it is applied to content.

_ParagraphStyle objects
class docx.styles.style._ParagraphStyle[source]
Bases: docx.styles.style._CharacterStyle

A paragraph style. A paragraph style provides both character formatting and paragraph formatting such as indentation and line-spacing.

base_style
Style object this style inherits from or None if this style is not based on another style.

builtin
Read-only. True if this style is a built-in style. False indicates it is a custom (user-defined) style. Note this value is based on the presence of a customStyle attribute in the XML, not on specific knowledge of which styles are built into Word.

delete()
Remove this style definition from the document. Note that calling this method does not remove or change the style applied to any document content. Content items having the deleted style will be rendered using the default style, as is any content with a style not defined in the document.

font
The Font object providing access to the character formatting properties for this style, such as font name and size.

hidden
True if display of this style in the style gallery and list of recommended styles is suppressed. False otherwise. In order to be shown in the style gallery, this value must be False and quick_style must be True.

locked
Read/write Boolean. True if this style is locked. A locked style does not appear in the styles panel or the style gallery and cannot be applied to document content. This behavior is only active when formatting protection is turned on for the document (via the Developer menu).

name
The UI name of this style.

next_paragraph_style
_ParagraphStyle object representing the style to be applied automatically to a new paragraph inserted after a paragraph of this style. Returns self if no next paragraph style is defined. Assigning None or self removes the setting such that new paragraphs are created using this same style.

paragraph_format
The ParagraphFormat object providing access to the paragraph formatting properties for this style such as indentation.

priority
The integer sort key governing display sequence of this style in the Word UI. None indicates no setting is defined, causing Word to use the default value of 0. Style name is used as a secondary sort key to resolve ordering of styles having the same priority value.

quick_style
True if this style should be displayed in the style gallery when hidden is False. Read/write Boolean.

unhide_when_used
True if an application should make this style visible the next time it is applied to content. False otherwise. Note that python-docx does not automatically unhide a style having True for this attribute when it is applied to content.

_TableStyle objects
class docx.styles.style._TableStyle[source]
Bases: docx.styles.style._ParagraphStyle

A table style. A table style provides character and paragraph formatting for its contents as well as special table formatting properties.

base_style
Style object this style inherits from or None if this style is not based on another style.

builtin
Read-only. True if this style is a built-in style. False indicates it is a custom (user-defined) style. Note this value is based on the presence of a customStyle attribute in the XML, not on specific knowledge of which styles are built into Word.

delete()
Remove this style definition from the document. Note that calling this method does not remove or change the style applied to any document content. Content items having the deleted style will be rendered using the default style, as is any content with a style not defined in the document.

font
The Font object providing access to the character formatting properties for this style, such as font name and size.

hidden
True if display of this style in the style gallery and list of recommended styles is suppressed. False otherwise. In order to be shown in the style gallery, this value must be False and quick_style must be True.

locked
Read/write Boolean. True if this style is locked. A locked style does not appear in the styles panel or the style gallery and cannot be applied to document content. This behavior is only active when formatting protection is turned on for the document (via the Developer menu).

name
The UI name of this style.

next_paragraph_style
_ParagraphStyle object representing the style to be applied automatically to a new paragraph inserted after a paragraph of this style. Returns self if no next paragraph style is defined. Assigning None or self removes the setting such that new paragraphs are created using this same style.

paragraph_format
The ParagraphFormat object providing access to the paragraph formatting properties for this style such as indentation.

priority
The integer sort key governing display sequence of this style in the Word UI. None indicates no setting is defined, causing Word to use the default value of 0. Style name is used as a secondary sort key to resolve ordering of styles having the same priority value.

quick_style
True if this style should be displayed in the style gallery when hidden is False. Read/write Boolean.

unhide_when_used
True if an application should make this style visible the next time it is applied to content. False otherwise. Note that python-docx does not automatically unhide a style having True for this attribute when it is applied to content.

_NumberingStyle objects
class docx.styles.style._NumberingStyle[source]
A numbering style. Not yet implemented.

LatentStyles objects
class docx.styles.latent.LatentStyles[source]
Provides access to the default behaviors for latent styles in this document and to the collection of _LatentStyle objects that define overrides of those defaults for a particular named latent style.

add_latent_style(name)[source]
Return a newly added _LatentStyle object to override the inherited defaults defined in this latent styles object for the built-in style having name.

default_priority
Integer between 0 and 99 inclusive specifying the default sort order for latent styles in style lists and the style gallery. None if no value is assigned, which causes Word to use the default value 99.

default_to_hidden
Boolean specifying whether the default behavior for latent styles is to be hidden. A hidden style does not appear in the recommended list or in the style gallery.

default_to_locked
Boolean specifying whether the default behavior for latent styles is to be locked. A locked style does not appear in the styles panel or the style gallery and cannot be applied to document content. This behavior is only active when formatting protection is turned on for the document (via the Developer menu).

default_to_quick_style
Boolean specifying whether the default behavior for latent styles is to appear in the style gallery when not hidden.

default_to_unhide_when_used
Boolean specifying whether the default behavior for latent styles is to be unhidden when first applied to content.

element
The lxml element proxied by this object.

load_count
Integer specifying the number of built-in styles to initialize to the defaults specified in this LatentStyles object. None if there is no setting in the XML (very uncommon). The default Word 2011 template sets this value to 276, accounting for the built-in styles in Word 2010.

_LatentStyle objects
class docx.styles.latent._LatentStyle[source]
Proxy for an w:lsdException element, which specifies display behaviors for a built-in style when no definition for that style is stored yet in the styles.xml part. The values in this element override the defaults specified in the parent w:latentStyles element.

delete()[source]
Remove this latent style definition such that the defaults defined in the containing LatentStyles object provide the effective value for each of its attributes. Attempting to access any attributes on this object after calling this method will raise AttributeError.

element
The lxml element proxied by this object.

hidden
Tri-state value specifying whether this latent style should appear in the recommended list. None indicates the effective value is inherited from the parent <w:latentStyles> element.

locked
Tri-state value specifying whether this latent styles is locked. A locked style does not appear in the styles panel or the style gallery and cannot be applied to document content. This behavior is only active when formatting protection is turned on for the document (via the Developer menu).

name
The name of the built-in style this exception applies to.

priority
The integer sort key for this latent style in the Word UI.

quick_style
Tri-state value specifying whether this latent style should appear in the Word styles gallery when not hidden. None indicates the effective value should be inherited from the default values in its parent LatentStyles object.

unhide_when_used
Tri-state value specifying whether this style should have its hidden attribute set False the next time the style is applied to content. None indicates the effective value should be inherited from the default specified by its parent LatentStyles object.

Paragraph objects
class docx.text.paragraph.Paragraph[source]
Proxy object wrapping <w:p> element.

add_run(text=None, style=None)[source]
Append a run to this paragraph containing text and having character style identified by style ID style. text can contain tab (\t) characters, which are converted to the appropriate XML form for a tab. text can also include newline (\n) or carriage return (\r) characters, each of which is converted to a line break.

alignment
A member of the WD_PARAGRAPH_ALIGNMENT enumeration specifying the justification setting for this paragraph. A value of None indicates the paragraph has no directly-applied alignment value and will inherit its alignment value from its style hierarchy. Assigning None to this property removes any directly-applied alignment value.

clear()[source]
Return this same paragraph after removing all its content. Paragraph-level formatting, such as style, is preserved.

insert_paragraph_before(text=None, style=None)[source]
Return a newly created paragraph, inserted directly before this paragraph. If text is supplied, the new paragraph contains that text in a single run. If style is provided, that style is assigned to the new paragraph.

paragraph_format
The ParagraphFormat object providing access to the formatting properties for this paragraph, such as line spacing and indentation.

runs
Sequence of Run instances corresponding to the <w:r> elements in this paragraph.

style
Read/Write. _ParagraphStyle object representing the style assigned to this paragraph. If no explicit style is assigned to this paragraph, its value is the default paragraph style for the document. A paragraph style name can be assigned in lieu of a paragraph style object. Assigning None removes any applied style, making its effective value the default paragraph style for the document.

text
String formed by concatenating the text of each run in the paragraph. Tabs and line breaks in the XML are mapped to \t and \n characters respectively.

Assigning text to this property causes all existing paragraph content to be replaced with a single run containing the assigned text. A \t character in the text is mapped to a <w:tab/> element and each \n or \r character is mapped to a line break. Paragraph-level formatting, such as style, is preserved. All run-level formatting, such as bold or italic, is removed.

ParagraphFormat objects
class docx.text.parfmt.ParagraphFormat[source]
Provides access to paragraph formatting such as justification, indentation, line spacing, space before and after, and widow/orphan control.

alignment
A member of the WD_PARAGRAPH_ALIGNMENT enumeration specifying the justification setting for this paragraph. A value of None indicates paragraph alignment is inherited from the style hierarchy.

first_line_indent
Length value specifying the relative difference in indentation for the first line of the paragraph. A positive value causes the first line to be indented. A negative value produces a hanging indent. None indicates first line indentation is inherited from the style hierarchy.

keep_together
True if the paragraph should be kept “in one piece” and not broken across a page boundary when the document is rendered. None indicates its effective value is inherited from the style hierarchy.

keep_with_next
True if the paragraph should be kept on the same page as the subsequent paragraph when the document is rendered. For example, this property could be used to keep a section heading on the same page as its first paragraph. None indicates its effective value is inherited from the style hierarchy.

left_indent
Length value specifying the space between the left margin and the left side of the paragraph. None indicates the left indent value is inherited from the style hierarchy. Use an Inches value object as a convenient way to apply indentation in units of inches.

line_spacing
float or Length value specifying the space between baselines in successive lines of the paragraph. A value of None indicates line spacing is inherited from the style hierarchy. A float value, e.g. 2.0 or 1.75, indicates spacing is applied in multiples of line heights. A Length value such as Pt(12) indicates spacing is a fixed height. The Pt value class is a convenient way to apply line spacing in units of points. Assigning None resets line spacing to inherit from the style hierarchy.

line_spacing_rule
A member of the WD_LINE_SPACING enumeration indicating how the value of line_spacing should be interpreted. Assigning any of the WD_LINE_SPACING members SINGLE, DOUBLE, or ONE_POINT_FIVE will cause the value of line_spacing to be updated to produce the corresponding line spacing.

page_break_before
True if the paragraph should appear at the top of the page following the prior paragraph. None indicates its effective value is inherited from the style hierarchy.

right_indent
Length value specifying the space between the right margin and the right side of the paragraph. None indicates the right indent value is inherited from the style hierarchy. Use a Cm value object as a convenient way to apply indentation in units of centimeters.

space_after
Length value specifying the spacing to appear between this paragraph and the subsequent paragraph. None indicates this value is inherited from the style hierarchy. Length objects provide convenience properties, such as pt and inches, that allow easy conversion to various length units.

space_before
Length value specifying the spacing to appear between this paragraph and the prior paragraph. None indicates this value is inherited from the style hierarchy. Length objects provide convenience properties, such as pt and cm, that allow easy conversion to various length units.

tab_stops
TabStops object providing access to the tab stops defined for this paragraph format.

widow_control
True if the first and last lines in the paragraph remain on the same page as the rest of the paragraph when Word repaginates the document. None indicates its effective value is inherited from the style hierarchy.

Run objects
class docx.text.run.Run[source]
Proxy object wrapping <w:r> element. Several of the properties on Run take a tri-state value, True, False, or None. True and False correspond to on and off respectively. None indicates the property is not specified directly on the run and its effective value is taken from the style hierarchy.

add_break(break_type=6)[source]
Add a break element of break_type to this run. break_type can take the values WD_BREAK.LINE, WD_BREAK.PAGE, and WD_BREAK.COLUMN where WD_BREAK is imported from docx.enum.text. break_type defaults to WD_BREAK.LINE.

add_picture(image_path_or_stream, width=None, height=None)[source]
Return an InlineShape instance containing the image identified by image_path_or_stream, added to the end of this run. image_path_or_stream can be a path (a string) or a file-like object containing a binary image. If neither width nor height is specified, the picture appears at its native size. If only one is specified, it is used to compute a scaling factor that is then applied to the unspecified dimension, preserving the aspect ratio of the image. The native size of the picture is calculated using the dots-per-inch (dpi) value specified in the image file, defaulting to 72 dpi if no value is specified, as is often the case.

add_tab()[source]
Add a <w:tab/> element at the end of the run, which Word interprets as a tab character.

add_text(text)[source]
Returns a newly appended _Text object (corresponding to a new <w:t> child element) to the run, containing text. Compare with the possibly more friendly approach of assigning text to the Run.text property.

bold
Read/write. Causes the text of the run to appear in bold.

clear()[source]
Return reference to this run after removing all its content. All run formatting is preserved.

font
The Font object providing access to the character formatting properties for this run, such as font name and size.

italic
Read/write tri-state value. When True, causes the text of the run to appear in italics.

style
Read/write. A _CharacterStyle object representing the character style applied to this run. The default character style for the document (often Default Character Font) is returned if the run has no directly-applied character style. Setting this property to None removes any directly-applied character style.

text
String formed by concatenating the text equivalent of each run content child element into a Python string. Each <w:t> element adds the text characters it contains. A <w:tab/> element adds a \t character. A <w:cr/> or <w:br> element each add a \n character. Note that a <w:br> element can indicate a page break or column break as well as a line break. All <w:br> elements translate to a single \n character regardless of their type. All other content child elements, such as <w:drawing>, are ignored.

Assigning text to this property has the reverse effect, translating each \t character to a <w:tab/> element and each \n or \r character to a <w:cr/> element. Any existing run content is replaced. Run formatting is preserved.

underline
The underline style for this Run, one of None, True, False, or a value from WD_UNDERLINE. A value of None indicates the run has no directly-applied underline value and so will inherit the underline value of its containing paragraph. Assigning None to this property removes any directly-applied underline value. A value of False indicates a directly-applied setting of no underline, overriding any inherited value. A value of True indicates single underline. The values from WD_UNDERLINE are used to specify other outline styles such as double, wavy, and dotted.

Font objects
class docx.text.run.Font[source]
Proxy object wrapping the parent of a <w:rPr> element and providing access to character properties such as font name, font size, bold, and subscript.

all_caps
Read/write. Causes text in this font to appear in capital letters.

bold
Read/write. Causes text in this font to appear in bold.

color
A ColorFormat object providing a way to get and set the text color for this font.

complex_script
Read/write tri-state value. When True, causes the characters in the run to be treated as complex script regardless of their Unicode values.

cs_bold
Read/write tri-state value. When True, causes the complex script characters in the run to be displayed in bold typeface.

cs_italic
Read/write tri-state value. When True, causes the complex script characters in the run to be displayed in italic typeface.

double_strike
Read/write tri-state value. When True, causes the text in the run to appear with double strikethrough.

emboss
Read/write tri-state value. When True, causes the text in the run to appear as if raised off the page in relief.

hidden
Read/write tri-state value. When True, causes the text in the run to be hidden from display, unless applications settings force hidden text to be shown.

highlight_color
A member of WD_COLOR_INDEX indicating the color of highlighting applied, or None if no highlighting is applied.

imprint
Read/write tri-state value. When True, causes the text in the run to appear as if pressed into the page.

italic
Read/write tri-state value. When True, causes the text of the run to appear in italics. None indicates the effective value is inherited from the style hierarchy.

math
Read/write tri-state value. When True, specifies this run contains WML that should be handled as though it was Office Open XML Math.

name
Get or set the typeface name for this Font instance, causing the text it controls to appear in the named font, if a matching font is found. None indicates the typeface is inherited from the style hierarchy.

no_proof
Read/write tri-state value. When True, specifies that the contents of this run should not report any errors when the document is scanned for spelling and grammar.

outline
Read/write tri-state value. When True causes the characters in the run to appear as if they have an outline, by drawing a one pixel wide border around the inside and outside borders of each character glyph.

rtl
Read/write tri-state value. When True causes the text in the run to have right-to-left characteristics.

shadow
Read/write tri-state value. When True causes the text in the run to appear as if each character has a shadow.

size
Read/write Length value or None, indicating the font height in English Metric Units (EMU). None indicates the font size should be inherited from the style hierarchy. Length is a subclass of int having properties for convenient conversion into points or other length units. The docx.shared.Pt class allows convenient specification of point values:

>'>>> font.size = Pt(24)  
>'>>> font.size  
304800  
>'>>> font.size.pt  
24.0  
small_caps  
Read/write tri-state value. When True causes the lowercase characters in the run to appear as capital letters two points smaller than the font size specified for the run.

snap_to_grid
Read/write tri-state value. When True causes the run to use the document grid characters per line settings defined in the docGrid element when laying out the characters in this run.

spec_vanish
Read/write tri-state value. When True, specifies that the given run shall always behave as if it is hidden, even when hidden text is being displayed in the current document. The property has a very narrow, specialized use related to the table of contents. Consult the spec (§17.3.2.36) for more details.

strike
Read/write tri-state value. When True causes the text in the run to appear with a single horizontal line through the center of the line.

subscript
Boolean indicating whether the characters in this Font appear as subscript. None indicates the subscript/subscript value is inherited from the style hierarchy.

superscript
Boolean indicating whether the characters in this Font appear as superscript. None indicates the subscript/superscript value is inherited from the style hierarchy.

underline
The underline style for this Font, one of None, True, False, or a value from WD_UNDERLINE. None indicates the font inherits its underline value from the style hierarchy. False indicates no underline. True indicates single underline. The values from WD_UNDERLINE are used to specify other outline styles such as double, wavy, and dotted.

web_hidden
Read/write tri-state value. When True, specifies that the contents of this run shall be hidden when the document is displayed in web page view.

TabStop objects
class docx.text.tabstops.TabStop[source]
An individual tab stop applying to a paragraph or style. Accessed using list semantics on its containing TabStops object.

alignment
A member of WD_TAB_ALIGNMENT specifying the alignment setting for this tab stop. Read/write.

leader
A member of WD_TAB_LEADER specifying a repeating character used as a “leader”, filling in the space spanned by this tab. Assigning None produces the same result as assigning WD_TAB_LEADER.SPACES. Read/write.

position
A Length object representing the distance of this tab stop from the inside edge of the paragraph. May be positive or negative. Read/write.

TabStops objects
class docx.text.tabstops.TabStops[source]
A sequence of TabStop objects providing access to the tab stops of a paragraph or paragraph style. Supports iteration, indexed access, del, and len(). It is accesed using the tab_stops property of ParagraphFormat; it is not intended to be constructed directly.

add_tab_stop(position, alignment=WD_TAB_ALIGNMENT.LEFT, leader=WD_TAB_LEADER.SPACES)[source]
Add a new tab stop at position, a Length object specifying the location of the tab stop relative to the paragraph edge. A negative position value is valid and appears in hanging indentation. Tab alignment defaults to left, but may be specified by passing a member of the WD_TAB_ALIGNMENT enumeration as alignment. An optional leader character can be specified by passing a member of the WD_TAB_LEADER enumeration as leader.

clear_all()[source]
Remove all custom tab stops.

Table objects
Table objects are constructed using the add_table() method on Document.

Table objects
class docx.table.Table(tbl, parent)[source]
Proxy class for a WordprocessingML <w:tbl> element.

add_column(width)[source]
Return a _Column object of width, newly added rightmost to the table.

add_row()[source]
Return a _Row instance, newly added bottom-most to the table.

alignment
Read/write. A member of WD_TABLE_ALIGNMENT or None, specifying the positioning of this table between the page margins. None if no setting is specified, causing the effective value to be inherited from the style hierarchy.

autofit
True if column widths can be automatically adjusted to improve the fit of cell contents. False if table layout is fixed. Column widths are adjusted in either case if total column width exceeds page width. Read/write boolean.

cell(row_idx, col_idx)[source]
Return _Cell instance correponding to table cell at row_idx, col_idx intersection, where (0, 0) is the top, left-most cell.

column_cells(column_idx)[source]
Sequence of cells in the column at column_idx in this table.

columns
_Columns instance representing the sequence of columns in this table.

row_cells(row_idx)[source]
Sequence of cells in the row at row_idx in this table.

rows
_Rows instance containing the sequence of rows in this table.

style
Read/write. A _TableStyle object representing the style applied to this table. The default table style for the document (often Normal Table) is returned if the table has no directly-applied style. Assigning None to this property removes any directly-applied table style causing it to inherit the default table style of the document. Note that the style name of a table style differs slightly from that displayed in the user interface; a hyphen, if it appears, must be removed. For example, Light Shading - Accent 1 becomes Light Shading Accent 1.

table_direction
A member of WD_TABLE_DIRECTION indicating the direction in which the table cells are ordered, e.g. WD_TABLE_DIRECTION.LTR. None indicates the value is inherited from the style hierarchy.

_Cell objects
class docx.table._Cell(tc, parent)[source]
Table cell

add_paragraph(text=u'', style=None)[source]
Return a paragraph newly added to the end of the content in this cell. If present, text is added to the paragraph in a single run. If specified, the paragraph style style is applied. If style is not specified or is None, the result is as though the ‘Normal’ style was applied. Note that the formatting of text in a cell can be influenced by the table style. text can contain tab (\t) characters, which are converted to the appropriate XML form for a tab. text can also include newline (\n) or carriage return (\r) characters, each of which is converted to a line break.

add_table(rows, cols)[source]
Return a table newly added to this cell after any existing cell content, having rows rows and cols columns. An empty paragraph is added after the table because Word requires a paragraph element as the last element in every cell.

merge(other_cell)[source]
Return a merged cell created by spanning the rectangular region having this cell and other_cell as diagonal corners. Raises InvalidSpanError if the cells do not define a rectangular region.

paragraphs
List of paragraphs in the cell. A table cell is required to contain at least one block-level element and end with a paragraph. By default, a new cell contains a single paragraph. Read-only

tables
List of tables in the cell, in the order they appear. Read-only.

text
The entire contents of this cell as a string of text. Assigning a string to this property replaces all existing content with a single paragraph containing the assigned text in a single run.

vertical_alignment
Member of WD_CELL_VERTICAL_ALIGNMENT or None.

A value of None indicates vertical alignment for this cell is inherited. Assigning None causes any explicitly defined vertical alignment to be removed, restoring inheritance.

width
The width of this cell in EMU, or None if no explicit width is set.

_Row objects
class docx.table._Row(tr, parent)[source]
Table row

cells
Sequence of _Cell instances corresponding to cells in this row.

height
Return a Length object representing the height of this cell, or None if no explicit height is set.

height_rule
Return the height rule of this cell as a member of the WD_ROW_HEIGHT_RULE enumeration, or None if no explicit height_rule is set.

table
Reference to the Table object this row belongs to.

_Column objects
class docx.table._Column(gridCol, parent)[source]
Table column

cells
Sequence of _Cell instances corresponding to cells in this column.

table
Reference to the Table object this column belongs to.

width
The width of this column in EMU, or None if no explicit width is set.

_Rows objects
class docx.table._Rows(tbl, parent)[source]
Sequence of _Row objects corresponding to the rows in a table. Supports len(), iteration, indexed access, and slicing.

table
Reference to the Table object this row collection belongs to.

_Columns objects
class docx.table._Columns(tbl, parent)[source]
Sequence of _Column instances corresponding to the columns in a table. Supports len(), iteration and indexed access.

table
Reference to the Table object this column collection belongs to.

API basics
The API for python-docx is designed to make doing simple things simple, while allowing more complex results to be achieved with a modest and incremental investment of understanding.

It’s possible to create a basic document using only a single object, the docx.api.Document object returned when opening a file. The methods on docx.api.Document allow block-level objects to be added to the end of the document. Block-level objects include paragraphs, inline pictures, and tables. Headings, bullets, and numbered lists are simply paragraphs with a particular style applied.

In this way, a document can be “written” from top to bottom, roughly like a person would if they knew exactly what they wanted to say This basic use case, where content is always added to the end of the document, is expected to account for perhaps 80% of actual use cases, so it’s a priority to make it as simple as possible without compromising the power of the overall API.

Inline objects
Each block-level method on docx.api.Document, such as add_paragraph(), returns the block-level object created. Often the reference is unneeded; but when inline objects must be created individually, you’ll need the block-item reference to do it.

… add example here as API solidifies …

## Understanding Styles¶
**Grasshopper:**  
*“Master, why doesn’t my paragraph appear with the style I specified?”*  
**Master:**  
*“You have come to the right page Grasshopper; read on …”*  
### What is a style in Word?  
Documents communicate better when like elements are formatted consistently. To achieve that consistency, professional document designers develop a style sheet which defines the document element types and specifies how each should be formatted. For example, perhaps body paragraphs are to be set in 9 pt Times Roman with a line height of 11 pt, justified flush left, ragged right. When these specifications are applied to each of the elements of the document, a consistent and polished look is achieved.

A style in Word is such a set of specifications that may be applied, all at once, to a document element. Word has paragraph styles, character styles, table styles, and numbering definitions. These are applied to a paragraph, a span of text, a table, and a list, respectively.

Experienced programmers will recognize styles as a level of indirection. The great thing about those is it allows you to define something once, then apply that definition many times. This saves the work of defining the same thing over an over; but more importantly it allows you to change the definition and have that change reflected in all the places you have applied it.

### Why doesn’t the style I applied show up?  
This is likely to show up quite a bit until I can add some fancier features to work around it, so here it is up top.    
1. When you’re working in Word, there are all these styles you can apply to things, pretty good looking ones that look all the better because you don’t have to make them yourself. Most folks never look further than the built-in styles.  
2. Although those styles show up in the UI, they’re not actually in the document you’re creating, at least not until you use it for the first time. That’s kind of a good thing. They take up room and there’s a lot of them. The file would get a little bloated if it contained all the style definitions you could use but haven’t.  
3. If you apply a style using python-docx that’s not defined in your file (in the styles.xml part if you’re curious), Word just ignores it. It doesn’t complain, it just doesn’t change how things are formatted. I’m sure there’s a good reason for this. But it can present as a bit of a puzzle if you don’t understand how Word works that way.  
4. When you use a style, Word adds it to the file. Once there, it stays. I imagine there’s a way to get rid of it, but you have to work at it. If you apply a style, delete the content you applied it to, and then save the document; the style definition stays in the saved file.  
All this adds up to the following: If you want to use a style in a document you create with python-docx, the document you start with must contain the style definition. Otherwise it just won’t work. It won’t raise an exception, it just won’t work.

If you use the “default” template document, it contains the styles listed below, most of the ones you’re likely to want if you’re not designing your own. If you’re using your own starting document, you need to use each of the styles you want at least once in it. You don’t have to keep the content, but you need to apply the style to something at least once before saving the document. Creating a one-word paragraph, applying five styles to it in succession and then deleting the paragraph works fine. That’s how I got the ones below into the default template :).

### Glossary
style definition
>A <w:style> element in the styles part of a document that explicitly defines the attributes of a style.

defined style  
>A style that is explicitly defined in a document. Contrast with latent style.

built-in style  
>One of the set of 276 pre-set styles built into Word, such as “Heading 1”. A built-in style can be either defined or latent. A built-in style that is not yet defined is known as a latent style. Both defined and latent built-in styles may appear as options in Word’s style panel and style gallery.

custom style  
>Also known as a user defined style, any style defined in a Word document that is not a built-in style. Note that a custom style cannot be a latent style.

latent style  
>A built-in style having no definition in a particular document is known as a latent style in that document. A latent style can appear as an option in the Word UI depending on the settings in the LatentStyles object for the document.

recommended style list  
>A list of styles that appears in the styles toolbox or panel when “Recommended” is selected from the “List:” dropdown box.

Style Gallery  
>The selection of example styles that appear in the ribbon of the Word UI and which may be applied by clicking on one of them.

### Identifying a style  
A style has three identifying properties, *name*, *style_id*, and *type*.

Each style’s **name** property is its stable, unique identifier for access purposes.

A style’s **style_id** is used internally to key a content object such as a paragraph to its style. However this value is generated automatically by Word and is not guaranteed to be stable across saves. In general, the style id is formed simply by removing spaces from the localized style name, however there are exceptions. Users of python-docx should generally avoid using the style id unless they are confident with the internals involved.

A style’s **type** is set at creation time and cannot be changed.

### Built-in styles  
Word comes with almost 300 so-called *built-in* styles like *Normal*, *Heading 1*, and *List Bullet*. Style definitions are stored in the *styles.xml* part of a .docx package, but built-in style definitions are stored in the Word application itself and are not written to *styles.xml* until they are actually used. This is a sensible strategy because they take up considerable room and would be largely redundant and useless overhead in every .docx file otherwise.

The fact that built-in styles are not written to the .docx package until used gives rise to the need for latent style definitions, explained below.

### Style Behavior  
In addition to collecting a set of formatting properties, a style has five properties that specify its behavior. This behavior is relatively simple, basically amounting to when and where the style appears in the Word or LibreOffice UI.

The key notion to understanding style behavior is the recommended list. In the style pane in Word, the user can select which list of styles they want to see. One of these is named *Recommended* and is known as the *recommended list*. All five behavior properties affect some aspect of the style’s appearance in this list and in the style gallery.

In brief, a style appears in the recommended list if its **hidden** property is **False** (the default). If a style is not hidden and its **quick_style** property is **True**, it also appears in the style gallery. If a hidden style’s **unhide_when_used** property is **True**, its hidden property is set **False** the first time it is used. Styles in the style lists and style gallery are sorted in **priority** order, then alphabetically for styles of the same priority. If a style’s **locked** property is **True** and formatting restrictions are turned on for the document, the style will not appear in any list or the style gallery and cannot be applied to content.

### Latent styles
The need to specify the UI behavior of built-in styles not defined in *styles.xml* gives rise to the need for latent style definitions. A latent style definition is basically a stub style definition that has at most the five behavior attributes in addition to the style name. Additional space is saved by defining defaults for each of the behavior attributes, so only those that differ from the default need be defined and styles that match all defaults need no latent style definition.

Latent style definitions are specified using the *w:latentStyles* and *w:lsdException* elements appearing in *styles.xml*.

A latent style definition is only required for a built-in style because only a built-in style can appear in the UI without a style definition in *styles.xml*.

### Style inheritance
A style can inherit properties from another style, somewhat similarly to how Cascading Style Sheets (CSS) works. Inheritance is specified using the *base_style* attribute. By basing one style on another, an inheritance hierarchy of arbitrary depth can be formed. A style having no base style inherits properties from the document defaults.

### Paragraph styles in default template
- Normal
- Body Text
- Body Text 2
- Body Text 3
- Caption
- Heading 1
- Heading 2
- Heading 3
- Heading 4
- Heading 5
- Heading 6
- Heading 7
- Heading 8
- Heading 9
- Intense Quote
- List
- List 2
- List 3
- List Bullet
- List Bullet 2
- List Bullet 3
- List Continue
- List Continue 2
- List Continue 3
- List Number
- List Number 2
- List Number 3
- List Paragraph
- Macro Text
- No Spacing
- Quote
- Subtitle
- TOCHeading
- Title
### Character styles in default template
- Body Text Char
- Body Text 2 Char
- Body Text 3 Char
- Book Title
- Default Paragraph Font
- Emphasis
- Heading 1 Char
- Heading 2 Char
- Heading 3 Char
- Heading 4 Char
- Heading 5 Char
- Heading 6 Char
- Heading 7 Char
- Heading 8 Char
- Heading 9 Char
- Intense Emphasis
- Intense Quote Char
- Intense Reference
- Macro Text Char
- Quote Char
- Strong
- Subtitle Char
- Subtle Emphasis
- Subtle Reference
- Title Char
### Table styles in default template
- Table Normal
- Colorful Grid
- Colorful Grid Accent 1
- Colorful Grid Accent 2
- Colorful Grid Accent 3
- Colorful Grid Accent 4
- Colorful Grid Accent 5
- Colorful Grid Accent 6
- Colorful List
- Colorful List Accent 1
- Colorful List Accent 2
- Colorful List Accent 3
- Colorful List Accent 4
- Colorful List Accent 5
- Colorful List Accent 6
- Colorful Shading
- Colorful Shading Accent 1
- Colorful Shading Accent 2
- Colorful Shading Accent 3
- Colorful Shading Accent 4
- Colorful Shading Accent 5
- Colorful Shading Accent 6
- Dark List
- Dark List Accent 1
- Dark List Accent 2
- Dark List Accent 3
- Dark List Accent 4
- Dark List Accent 5
- Dark List Accent 6
- Light Grid
- Light Grid Accent 1
- Light Grid Accent 2
- Light Grid Accent 3
- Light Grid Accent 4
- Light Grid Accent 5
- Light Grid Accent 6
- Light List
- Light List Accent 1
- Light List Accent 2
- Light List Accent 3
- Light List Accent 4
- Light List Accent 5
- Light List Accent 6
- Light Shading
- Light Shading Accent 1
- Light Shading Accent 2
- Light Shading Accent 3
- Light Shading Accent 4
- Light Shading Accent 5
- Light Shading Accent 6
- Medium Grid 1
- Medium Grid 1 Accent 1
- Medium Grid 1 Accent 2
- Medium Grid 1 Accent 3
- Medium Grid 1 Accent 4
- Medium Grid 1 Accent 5
- Medium Grid 1 Accent 6
- Medium Grid 2
- Medium Grid 2 Accent 1
- Medium Grid 2 Accent 2
- Medium Grid 2 Accent 3
- Medium Grid 2 Accent 4
- Medium Grid 2 Accent 5
- Medium Grid 2 Accent 6
- Medium Grid 3
- Medium Grid 3 Accent 1
- Medium Grid 3 Accent 2
- Medium Grid 3 Accent 3
- Medium Grid 3 Accent 4
- Medium Grid 3 Accent 5
- Medium Grid 3 Accent 6
- Medium List 1
- Medium List 1 Accent 1
- Medium List 1 Accent 2
- Medium List 1 Accent 3
- Medium List 1 Accent 4
- Medium List 1 Accent 5
- Medium List 1 Accent 6
- Medium List 2
- Medium List 2 Accent 1
- Medium List 2 Accent 2
- Medium List 2 Accent 3
- Medium List 2 Accent 4
- Medium List 2 Accent 5
- Medium List 2 Accent 6
- Medium Shading 1
- Medium Shading 1 Accent 1
- Medium Shading 1 Accent 2
- Medium Shading 1 Accent 3
- Medium Shading 1 Accent 4
- Medium Shading 1 Accent 5
- Medium Shading 1 Accent 6
- Medium Shading 2
- Medium Shading 2 Accent 1
- Medium Shading 2 Accent 2
- Medium Shading 2 Accent 3
- Medium Shading 2 Accent 4
- Medium Shading 2 Accent 5
- Medium Shading 2 Accent 6
- Table Grid

## Working with Styles
This page uses concepts developed in the prior page without introduction. If a term is unfamiliar, consult the prior page ~~Understanding Styles~~ for a definition.

Access a style
Styles are accessed using the ~~**Document.styles**~~ attribute:

>'>>> document = Document()'   
>'>>> styles = document.styles'  
>'>>> styles'  
<docx.styles.styles.Styles object at 0x10a7c4f50>  

The ~~**Styles**~~ object provides dictionary-style access to defined styles by name:

>'>>> styles['Normal']'  
<docx.styles.style._ParagraphStyle object at <0x10a7c4f6b>

>Note  
>Built-in styles are stored in a WordprocessingML file using their English name, e.g. ‘Heading 1’, even though users working on a localized version of Word will see native language names in the UI, e.g. ‘Kop 1’. Because python-docx operates on the WordprocessingML file, style lookups must use the English name. A document available on this external site allows you to create a mapping between local language names and English style names: ~~http://www.thedoctools.com/index.php?show=mt_create_style_name_list~~  
User-defined styles, also known as custom styles, are not localized and are accessed with the name exactly as it appears in the Word UI.

The **Styles** object is also iterable. By using the identification properties on BaseStyle, various subsets of the defined styles can be generated. For example, this code will produce a list of the defined paragraph styles:

>'>>> from docx.enum.style import WD_STYLE_TYPE  
>'>>> styles = document.styles  
>'>>> paragraph_styles = [  
...     s for s in styles if s.type == WD_STYLE_TYPE.PARAGRAPH
... ]  
>'>>> for style in paragraph_styles:  
...     print(style.name)  
...  
Normal  
Body Text  
List Bullet  
Apply a style  
The Paragraph, Run, and Table objects each have a style attribute. Assigning a style object to this attribute applies that style:

>'>>> document = Document()
>'>>> paragraph = document.add_paragraph()
>'>>> paragraph.style
<docx.styles.style._ParagraphStyle object at <0x11a7c4c50>
>'>>> paragraph.style.name
'Normal'
>'>>> paragraph.style = document.styles['Heading 1']
>'>>> paragraph.style.name
'Heading 1'
A style name can also be assigned directly, in which case python-docx will do the lookup for you:

>'>>> paragraph.style = 'List Bullet'
>'>>> paragraph.style
<docx.styles.style._ParagraphStyle object at <0x10a7c4f84>
>'>>> paragraph.style.name
'List Bullet'
A style can also be applied at creation time using either the style object or its name:

>'>>> paragraph = document.add_paragraph(style='Body Text')
>'>>> paragraph.style.name
'Body Text'
>'>>> body_text_style = document.styles['Body Text']
>'>>> paragraph = document.add_paragraph(style=body_text_style)
>'>>> paragraph.style.name
'Body Text'
Add or delete a style
A new style can be added to the document by specifying a unique name and a style type:

>'>>> from docx.enum.style import WD_STYLE_TYPE
>'>>> styles = document.styles
>'>>> style = styles.add_style('Citation', WD_STYLE_TYPE.PARAGRAPH)
>'>>> style.name
'Citation'
>'>>> style.type
PARAGRAPH (1)
Use the base_style property to specify a style the new style should inherit formatting settings from:

>'>>> style.base_style
None
>'>>> style.base_style = styles['Normal']
>'>>> style.base_style
<docx.styles.style._ParagraphStyle object at 0x10a7a9550>
>'>>> style.base_style.name
'Normal'
A style can be removed from the document simply by calling its delete() method:

>'>>> styles = document.styles
>'>>> len(styles)
10
>'>>> styles['Citation'].delete()
>'>>> len(styles)
9
Note

The Style.delete() method removes the style’s definition from the document. It does not affect content in the document to which that style is applied. Content having a style not defined in the document is rendered using the default style for that content object, e.g. ‘Normal’ in the case of a paragraph.

Define character formatting
Character, paragraph, and table styles can all specify character formatting to be applied to content with that style. All the character formatting that can be applied directly to text can be specified in a style. Examples include font typeface and size, bold, italic, and underline.

Each of these three style types have a font attribute providing access to a Font object. A style’s Font object provides properties for getting and setting the character formatting for that style.

Several examples are provided here. For a complete set of the available properties, see the Font API documentation.

The font for a style can be accessed like this:

>'>>> from docx import Document
>'>>> document = Document()
>'>>> style = document.styles['Normal']
>'>>> font = style.font
Typeface and size are set like this:

>'>>> from docx.shared import Pt
>'>>> font.name = 'Calibri'
>'>>> font.size = Pt(12)
Many font properties are tri-state, meaning they can take the values True, False, and None. True means the property is “on”, False means it is “off”. Conceptually, the None value means “inherit”. Because a style exists in an inheritance hierarchy, it is important to have the ability to specify a property at the right place in the hierarchy, generally as far up the hierarchy as possible. For example, if all headings should be in the Arial typeface, it makes more sense to set that property on the Heading 1 style and have Heading 2 inherit from Heading 1.

Bold and italic are tri-state properties, as are all-caps, strikethrough, superscript, and many others. See the Font API documentation for a full list:

>'>>> font.bold, font.italic
(None, None)
>'>>> font.italic = True
>'>>> font.italic
True
>'>>> font.italic = False
>'>>> font.italic
False
>'>>> font.italic = None
>'>>> font.italic
None
Underline is a bit of a special case. It is a hybrid of a tri-state property and an enumerated value property. True means single underline, by far the most common. False means no underline, but more often None is the right choice if no underlining is wanted since it is rare to inherit it from a base style. The other forms of underlining, such as double or dashed, are specified with a member of the WD_UNDERLINE enumeration:

>'>>> font.underline
None
>'>>> font.underline = True
>'>>> # or perhaps
>'>>> font.underline = WD_UNDERLINE.DOT_DASH
Define paragraph formatting
Both a paragraph style and a table style allow paragraph formatting to be specified. These styles provide access to a ParagraphFormat object via their paragraph_format property.

Paragraph formatting includes layout behaviors such as justification, indentation, space before and after, page break before, and widow/orphan control. For a complete list of the available properties, consult the API documentation page for the ParagraphFormat object.

Here’s an example of how you would create a paragraph style having hanging indentation of 1/4 inch, 12 points spacing above, and widow/orphan control:

>'>>> from docx.enum.style import WD_STYLE_TYPE
>'>>> from docx.shared import Inches, Pt
>'>>> document = Document()
>'>>> style = document.styles.add_style('Indent', WD_STYLE_TYPE.PARAGRAPH)
>'>>> paragraph_format = style.paragraph_format
>'>>> paragraph_format.left_indent = Inches(0.25)
>'>>> paragraph_format.first_line_indent = Inches(-0.25)
>'>>> paragraph_format.space_before = Pt(12)
>'>>> paragraph_format.widow_control = True
Use paragraph-specific style properties
A paragraph style has a next_paragraph_style property that specifies the style to be applied to new paragraphs inserted after a paragraph of that style. This is most useful when the style would normally appear only once in a sequence, such as a heading. In that case, the paragraph style can automatically be set back to a body style after completing the heading.

In the most common case (body paragraphs), subsequent paragraphs should receive the same style as the current paragraph. The default handles this case well by applying the same style if a next paragraph style is not specified.

Here’s an example of how you would change the next paragraph style of the Heading 1 style to Body Text:

>'>>> from docx import Document
>'>>> document = Document()
>'>>> styles = document.styles

>'>>> styles['Heading 1'].next_paragraph_style = styles['Body Text']
The default behavior can be restored by assigning None or the style itself:

>'>>> heading_1_style = styles['Heading 1']
>'>>> heading_1_style.next_paragraph_style.name
'Body Text'

>'>>> heading_1_style.next_paragraph_style = heading_1_style
>'>>> heading_1_style.next_paragraph_style.name
'Heading 1'

>'>>> heading_1_style.next_paragraph_style = None
>'>>> heading_1_style.next_paragraph_style.name
'Heading 1'
Control how a style appears in the Word UI
The properties of a style fall into two categories, behavioral properties and formatting properties. Its behavioral properties control when and where the style appears in the Word UI. Its formatting properties determine the formatting of content to which the style is applied, such as the size of the font and its paragraph indentation.

There are five behavioral properties of a style:

hidden
unhide_when_used
priority
quick_style
locked
See the Style Behavior section in Understanding Styles for a description of how these behavioral properties interact to determine when and where a style appears in the Word UI.

The priority property takes an integer value. The other four style behavior properties are tri-state, meaning they can take the value True (on), False (off), or None (inherit).

Display a style in the style gallery
The following code will cause the ‘Body Text’ paragraph style to appear first in the style gallery:

>'>>> from docx import Document
>'>>> document = Document()
>'>>> style = document.styles['Body Text']

>'>>> style.hidden = False
>'>>> style.quick_style = True
>'>>> style.priorty = 1
Remove a style from the style gallery
This code will remove the ‘Normal’ paragraph style from the style gallery, but allow it to remain in the recommended list:

>'>>> style = document.styles['Normal']

>'>>> style.hidden = False
>'>>> style.quick_style = False
Working with Latent Styles
See the Built-in styles and Latent styles sections in Understanding Styles for a description of how latent styles define the behavioral properties of built-in styles that are not yet defined in the styles.xml part of a .docx file.

Access the latent styles in a document
The latent styles in a document are accessed from the styles object:

>'>>> document = Document()
>'>>> latent_styles = document.styles.latent_styles
A LatentStyles object supports len(), iteration, and dictionary-style access by style name:

>'>>> len(latent_styles)
161

>'>>> latent_style_names = [ls.name for ls in latent_styles]
>'>>> latent_style_names
['Normal', 'Heading 1', 'Heading 2', ... 'TOC Heading']

>'>>> latent_quote = latent_styles['Quote']
>'>>> latent_quote
<docx.styles.latent.LatentStyle object at 0x10a7c4f50>
>'>>> latent_quote.priority
29
Change latent style defaults
The LatentStyles object also provides access to the default behavioral properties for built-in styles in the current document. These defaults provide the value for any undefined attributes of the _LatentStyle definitions and to all behavioral properties of built-in styles having no explicit latent style definition. See the API documentation for the LatentStyles object for the complete set of available properties:

>'>>> latent_styles.default_to_locked
False
>'>>> latent_styles.default_to_locked = True
>'>>> latent_styles.default_to_locked
True
Add a latent style definition
A new latent style can be added using the add_latent_style() method on LatentStyles. This code adds a new latent style for the builtin style ‘List Bullet’, setting it to appear in the style gallery:

>'>>> latent_style = latent_styles['List Bullet']
KeyError: no latent style with name 'List Bullet'
>'>>> latent_style = latent_styles.add_latent_style('List Bullet')
>'>>> latent_style.hidden = False
>'>>> latent_style.priority = 2
>'>>> latent_style.quick_style = True
Delete a latent style definition
A latent style definition can be deleted by calling its delete() method:

>'>>> latent_styles['Light Grid']
<docx.styles.latent.LatentStyle object at 0x10a7c4f50>
>'>>> latent_styles['Light Grid'].delete()
>'>>> latent_styles['Light Grid']
KeyError: no latent style with name 'Light Grid'