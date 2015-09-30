module Axlsx

  # XML Encoding
  ENCODING = "UTF-8".freeze

  # spreadsheetML namespace
  XML_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main".freeze

  # content-types namespace
  XML_NS_T = "http://schemas.openxmlformats.org/package/2006/content-types".freeze

  # extended-properties namespace
  APP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties".freeze

  # doc props namespace
  APP_NS_VT = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes".freeze

  # core properties namespace
  CORE_NS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties".freeze

  # dc elements (core) namespace
  CORE_NS_DC = "http://purl.org/dc/elements/1.1/".freeze

  # dcmit (core) namespcace
  CORE_NS_DCMIT = "http://purl.org/dc/dcmitype/".freeze

  # dc terms namespace
  CORE_NS_DCT = "http://purl.org/dc/terms/".freeze

  # xml schema namespace
  CORE_NS_XSI = "http://www.w3.org/2001/XMLSchema-instance".freeze

  # Digital signature namespace
  DIGITAL_SIGNATURE_NS = "http://schemas.openxmlformats.org/package/2006/digital-signature".freeze

  # spreadsheet drawing namespace
  XML_NS_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing".freeze

  # drawing namespace
  XML_NS_A  = "http://schemas.openxmlformats.org/drawingml/2006/main".freeze

  # chart namespace
  XML_NS_C  = "http://schemas.openxmlformats.org/drawingml/2006/chart".freeze

  # relationships namespace
  XML_NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships".freeze

  # relationships name space
  RELS_R = "http://schemas.openxmlformats.org/package/2006/relationships".freeze

  # table rels namespace
  TABLE_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table".freeze

  # pivot table rels namespace
  PIVOT_TABLE_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable".freeze

  # pivot table cache definition namespace
  PIVOT_TABLE_CACHE_DEFINITION_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition".freeze

  # workbook rels namespace
  WORKBOOK_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument".freeze

  # worksheet rels namespace
  WORKSHEET_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet".freeze

  # app rels namespace
  APP_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties".freeze

  # core rels namespace
  CORE_R = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties".freeze

  # digital signature rels namespace
  DIGITAL_SIGNATURE_R = "http://schemas.openxmlformats.org/package/2006/relationships/digital- signature/signature".freeze

  # styles rels namespace
  STYLES_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles".freeze

  # shared strings namespace
  SHARED_STRINGS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings".freeze

  # drawing rels namespace
  DRAWING_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing".freeze

  # chart rels namespace
  CHART_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart".freeze

  # image rels namespace
  IMAGE_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image".freeze

  # hyperlink rels namespace
  HYPERLINK_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink".freeze

  # comment rels namespace
  COMMENT_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments".freeze

  # comment relation for nil target
  COMMENT_R_NULL = "http://purl.oclc.org/ooxml/officeDocument/relationships/comments".freeze

  #vml drawing relation namespace
  VML_DRAWING_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing'

  # VML Drawing content type
  VML_DRAWING_CT = "application/vnd.openxmlformats-officedocument.vmlDrawing".freeze

  # table content type
  TABLE_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml".freeze

  # pivot table content type
  PIVOT_TABLE_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml".freeze

  # pivot table cache definition content type
  PIVOT_TABLE_CACHE_DEFINITION_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml".freeze

  # workbook content type
  WORKBOOK_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml".freeze

  # app content type
  APP_CT = "application/vnd.openxmlformats-officedocument.extended-properties+xml".freeze

  # rels content type
  RELS_CT = "application/vnd.openxmlformats-package.relationships+xml".freeze

  # styles content type
  STYLES_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml".freeze

  # xml content type
  XML_CT = "application/xml".freeze

  # worksheet content type
  WORKSHEET_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml".freeze

  # shared strings content type
  SHARED_STRINGS_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml".freeze

  # core content type
  CORE_CT = "application/vnd.openxmlformats-package.core-properties+xml".freeze

  # digital signature xml content type
  DIGITAL_SIGNATURE_XML_CT = "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml".freeze

  # digital signature origin content type
  DIGITAL_SIGNATURE_ORIGIN_CT = "application/vnd.openxmlformats-package.digital-signature-origin".freeze

  # digital signature certificate content type
  DIGITAL_SIGNATURE_CERTIFICATE_CT = "application/vnd.openxmlformats-package.digital-signature-certificate".freeze

  # chart content type
  CHART_CT = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml".freeze

  # comments content type
  COMMENT_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml".freeze

  # jpeg content type
  JPEG_CT = "image/jpeg".freeze

  # gif content type
  GIF_CT = "image/gif".freeze

  # png content type
  PNG_CT = "image/png".freeze

  #drawing content type
  DRAWING_CT = "application/vnd.openxmlformats-officedocument.drawing+xml".freeze


  # xml content type extensions
  XML_EX = "xml".freeze

  # jpeg extension
  JPEG_EX = "jpeg".freeze

  # gif extension
  GIF_EX = "gif".freeze

  # png extension
  PNG_EX = "png".freeze

  # rels content type extension
  RELS_EX = "rels".freeze

  # workbook part
  WORKBOOK_PN = "xl/workbook.xml".freeze

  # styles part
  STYLES_PN = "styles.xml".freeze

  # shared_strings  part
  SHARED_STRINGS_PN = "sharedStrings.xml".freeze

  # app part
  APP_PN = "docProps/app.xml".freeze

  # core part
  CORE_PN = "docProps/core.xml".freeze

  # content types part
  CONTENT_TYPES_PN = "[Content_Types].xml".freeze

  # rels part
  RELS_PN = "_rels/.rels".freeze

  # workbook rels part
  WORKBOOK_RELS_PN = "xl/_rels/workbook.xml.rels".freeze

  # worksheet part
  WORKSHEET_PN = "worksheets/sheet%d.xml".freeze

  # worksheet rels part
  WORKSHEET_RELS_PN = "worksheets/_rels/sheet%d.xml.rels".freeze

  # drawing part
  DRAWING_PN = "drawings/drawing%d.xml".freeze

  # drawing rels part
  DRAWING_RELS_PN = "drawings/_rels/drawing%d.xml.rels".freeze

  # vml drawing part
  VML_DRAWING_PN = "drawings/vmlDrawing%d.vml".freeze

  # drawing part
  TABLE_PN = "tables/table%d.xml".freeze

  # pivot table parts
  PIVOT_TABLE_PN = "pivotTables/pivotTable%d.xml".freeze

  # pivot table cache definition part name
  PIVOT_TABLE_CACHE_DEFINITION_PN = "pivotCache/pivotCacheDefinition%d.xml".freeze

  # pivot table rels parts
  PIVOT_TABLE_RELS_PN = "pivotTables/_rels/pivotTable%d.xml.rels".freeze

  # chart part
  CHART_PN = "charts/chart%d.xml".freeze

  # chart part
  IMAGE_PN = "media/image%d.%s".freeze

  # comment part
  COMMENT_PN = "comments%d.xml".freeze

  # location of schema files for validation
  SCHEMA_BASE = (File.dirname(__FILE__)+'/../../schema/').freeze

  # App validation schema
  APP_XSD = (SCHEMA_BASE + "shared-documentPropertiesExtended.xsd").freeze

  # core validation schema
  CORE_XSD = (SCHEMA_BASE + "opc-coreProperties.xsd").freeze

  # content types validation schema
  CONTENT_TYPES_XSD = (SCHEMA_BASE + "opc-contentTypes.xsd").freeze

  # rels validation schema
  RELS_XSD = (SCHEMA_BASE + "opc-relationships.xsd").freeze

  # spreadsheetML validation schema
  SML_XSD = (SCHEMA_BASE + "sml.xsd").freeze

  # drawing validation schema
  DRAWING_XSD = (SCHEMA_BASE + "dml-spreadsheetDrawing.xsd").freeze

  # number format id for pecentage formatting using the default formatting id.
  NUM_FMT_PERCENT = 9

  # number format id for date format like 2011/11/13
  NUM_FMT_YYYYMMDD = 100

  # number format id for time format the creates 2011/11/13 12:23:10
  NUM_FMT_YYYYMMDDHHMMSS = 101

  # cellXfs id for thin borders around the cell
  STYLE_THIN_BORDER = 1

  # cellXfs id for default date styling
  STYLE_DATE = 2

  # error messages RestrictionValidor
  ERR_RESTRICTION = "Invalid Data: %s. %s must be one of %s.".freeze

  # error message DataTypeValidator
  ERR_TYPE = "Invalid Data %s for %s. must be %s.".freeze

  # error message for RegexValidator
  ERR_REGEX = "Invalid Data. %s does not match %s.".freeze

  # error message for RangeValidator
  ERR_RANGE = "Invalid Data. %s must be between %s and %s, (inclusive:%s) you gave: %s".freeze

  # error message for sheets that use a name which is longer than 31 bytes
  ERR_SHEET_NAME_TOO_LONG = "Your worksheet name '%s' is too long. Worksheet names must be 31 characters (bytes) or less".freeze

  # error message for sheets that use a name which include invalid characters
  ERR_SHEET_NAME_CHARACTER_FORBIDDEN = "Your worksheet name '%s' contains a character which is not allowed by MS Excel and will cause repair warnings. Please change the name of your sheet.".freeze

  # error message for duplicate sheet names
  ERR_DUPLICATE_SHEET_NAME = "There is already a worksheet in this workbook named '%s'. Please use a unique name".freeze

  # error message when user does not provide color and or style options for border in Style#add_sytle
  ERR_INVALID_BORDER_OPTIONS = "border hash must include both style and color. e.g. :border => { :color => 'FF000000', :style => :thin }. You provided: %s".freeze

  # error message for invalid border id reference
  ERR_INVALID_BORDER_ID = "The border id you specified (%s) does not exist. Please add a border with Style#add_style before referencing its index.".freeze

  # error message for invalid angles
  ERR_ANGLE = "Angles must be a value between -90 and 90. You provided: %s".freeze

  # error message for non 'integerish' value
  ERR_INTEGERISH = "You value must be, or be castable via to_i, an Integer. You provided %s".freeze

  # Regex to match forbidden control characters
  # The following will be automatically stripped from worksheets.
  #
  # x00 Null
  # x01 Start Of Heading
  # x02 Start Of Text
  # x03End Of Text
  # x04 End Of Transmission
  # x05 Enquiry
  # x06 Acknowledge
  # x07 Bell
  # x08 Backspace
  # x0B Line Tabulation
  # x0C Form Feed
  # x0E Shift Out
  # x0F Shift In
  # x10 Data Link Escape
  # x11 Device Control One
  # x12 Device Control Two
  # x13 Device Control Three
  # x14 Device Control Four
  # x15 Negative Acknowledge
  # x16 Synchronous Idle
  # x17 End Of Transmission Block
  # x18 Cancel
  # x19 End Of Medium
  # x1A Substitute
  # x1B Escape
  # x1C Information Separator Four
  # x1D Information Separator Three
  # x1E Information Separator Two
  # x1F Information Separator One
  #
  # The following are not dealt with.
  # If you have this in your data, expect excel to blow up!
  #
  # x7F	Delete
  # x80	Control 0080
  # x81	Control 0081
  # x82	Break Permitted Here
  # x83	No Break Here
  # x84	Control 0084
  # x85	Next Line (Nel)
  # x86	Start Of Selected Area
  # x87	End Of Selected Area
  # x88	Character Tabulation Set
  # x89	Character Tabulation With Justification
  # x8A	Line Tabulation Set
  # x8B	Partial Line Forward
  # x8C	Partial Line Backward
  # x8D	Reverse Line Feed
  # x8E	Single Shift Two
  # x8F	Single Shift Three
  # x90	Device Control String
  # x91	Private Use One
  # x92	Private Use Two
  # x93	Set Transmit State
  # x94	Cancel Character
  # x95	Message Waiting
  # x96	Start Of Guarded Area
  # x97	End Of Guarded Area
  # x98	Start Of String
  # x99	Control 0099
  # x9A	Single Character Introducer
  # x9B	Control Sequence Introducer
  # x9C	String Terminator
  # x9D	Operating System Command
  # x9E	Privacy Message
  # x9F	Application Program Command
  #
  # The following are allowed:
  #
  # x0A Line Feed (Lf)
  # x0D Carriage Return (Cr)
  # x09 Character Tabulation
  # @see http://www.codetable.net/asciikeycodes
  pattern = "\x0-\x08\x0B\x0C\x0E-\x1F"
  pattern = pattern.respond_to?(:encode) ? pattern.encode('UTF-8') : pattern
  
  # The regular expression used to remove control characters from worksheets
  CONTROL_CHARS = pattern.freeze
  
  ISO_8601_REGEX = /\A(-?(?:[1-9][0-9]*)?[0-9]{4})-(1[0-2]|0[1-9])-(3[0-1]|0[1-9]|[1-2][0-9])T(2[0-3]|[0-1][0-9]):([0-5][0-9]):([0-5][0-9])(\.[0-9]+)?(Z|[+-](?:2[0-3]|[0-1][0-9]):[0-5][0-9])?\Z/.freeze
  
  FLOAT_REGEX = /\A[-+]?[0-9]*\.?[0-9]+([eE][-+]?[0-9]+)?\Z/.freeze
  
  NUMERIC_REGEX = /\A[+-]?\d+?\Z/.freeze
end
