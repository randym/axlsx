module Axlsx

  # XML Encoding
  ENCODING = "UTF-8"

  # spreadsheetML namespace
  XML_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

  # content-types namespace
  XML_NS_T = "http://schemas.openxmlformats.org/package/2006/content-types"

  # extended-properties namespace
  APP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"

  # doc props namespace
  APP_NS_VT = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

  # core properties namespace
  CORE_NS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"

  # dc elements (core) namespace
  CORE_NS_DC = "http://purl.org/dc/elements/1.1/"

  # dcmit (core) namespcace
  CORE_NS_DCMIT = "http://purl.org/dc/dcmitype/"

  # dc terms namespace
  CORE_NS_DCT = "http://purl.org/dc/terms/"

  # xml schema namespace
  CORE_NS_XSI = "http://www.w3.org/2001/XMLSchema-instance"

  # Digital signature namespace
  DIGITAL_SIGNATURE_NS = "http://schemas.openxmlformats.org/package/2006/digital-signature"

  # spreadsheet drawing namespace
  XML_NS_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"

  # drawing namespace
  XML_NS_A  = "http://schemas.openxmlformats.org/drawingml/2006/main"

  # chart namespace
  XML_NS_C  = "http://schemas.openxmlformats.org/drawingml/2006/chart"

  # relationships namespace
  XML_NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

  # relationships name space
  RELS_R = "http://schemas.openxmlformats.org/package/2006/relationships"

  # table rels namespace
  TABLE_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table"

  # pivot table rels namespace
  PIVOT_TABLE_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable"

  # pivot table cache definition namespace
  PIVOT_TABLE_CACHE_DEFINITION_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition"

  # workbook rels namespace
  WORKBOOK_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"

  # worksheet rels namespace
  WORKSHEET_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"

  # app rels namespace
  APP_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"

  # core rels namespace
  CORE_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata/core-properties"

  # digital signature rels namespace
  DIGITAL_SIGNATURE_R = "http://schemas.openxmlformats.org/package/2006/relationships/digital- signature/signature"

  # styles rels namespace
  STYLES_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"

  # shared strings namespace
  SHARED_STRINGS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"

  # drawing rels namespace
  DRAWING_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"

  # chart rels namespace
  CHART_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"

  # image rels namespace
  IMAGE_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"

  # hyperlink rels namespace
  HYPERLINK_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"

  # comment rels namespace
  COMMENT_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"

  # comment relation for nil target
  COMMENT_R_NULL = "http://purl.oclc.org/ooxml/officeDocument/relationships/comments"

  #vml drawing relation namespace
  VML_DRAWING_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing'

  # VML Drawing content type
  VML_DRAWING_CT = "application/vnd.openxmlformats-officedocument.vmlDrawing"

  # table content type
  TABLE_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"

  # pivot table content type
  PIVOT_TABLE_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"

  # pivot table cache definition content type
  PIVOT_TABLE_CACHE_DEFINITION_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"

  # workbook content type
  WORKBOOK_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"

  # app content type
  APP_CT = "application/vnd.openxmlformats-officedocument.extended-properties+xml"

  # rels content type
  RELS_CT = "application/vnd.openxmlformats-package.relationships+xml"

  # styles content type
  STYLES_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"

  # xml content type
  XML_CT = "application/xml"

  # worksheet content type
  WORKSHEET_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"

  # shared strings content type
  SHARED_STRINGS_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"

  # core content type
  CORE_CT = "application/vnd.openxmlformats-package.core-properties+xml"

  # digital signature xml content type
  DIGITAL_SIGNATURE_XML_CT = "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml"

  # digital signature origin content type
  DIGITAL_SIGNATURE_ORIGIN_CT = "application/vnd.openxmlformats-package.digital-signature-origin"

  # digital signature certificate content type
  DIGITAL_SIGNATURE_CERTIFICATE_CT = "application/vnd.openxmlformats-package.digital-signature-certificate"

  # chart content type
  CHART_CT = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"

  # comments content type
  COMMENT_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"

  # jpeg content type
  JPEG_CT = "image/jpeg"

  # gif content type
  GIF_CT = "image/gif"

  # png content type
  PNG_CT = "image/png"

  #drawing content type
  DRAWING_CT = "application/vnd.openxmlformats-officedocument.drawing+xml"


  # xml content type extensions
  XML_EX = "xml"

  # jpeg extension
  JPEG_EX = "jpeg"

  # gif extension
  GIF_EX = "gif"

  # png extension
  PNG_EX = "png"

  # rels content type extension
  RELS_EX = "rels"

  # workbook part
  WORKBOOK_PN = "xl/workbook.xml"

  # styles part
  STYLES_PN = "styles.xml"

  # shared_strings  part
  SHARED_STRINGS_PN = "sharedStrings.xml"

  # app part
  APP_PN = "docProps/app.xml"

  # core part
  CORE_PN = "docProps/core.xml"

  # content types part
  CONTENT_TYPES_PN = "[Content_Types].xml"

  # rels part
  RELS_PN = "_rels/.rels"

  # workbook rels part
  WORKBOOK_RELS_PN = "xl/_rels/workbook.xml.rels"

  # worksheet part
  WORKSHEET_PN = "worksheets/sheet%d.xml"

  # worksheet rels part
  WORKSHEET_RELS_PN = "worksheets/_rels/sheet%d.xml.rels"

  # drawing part
  DRAWING_PN = "drawings/drawing%d.xml"

  # drawing rels part
  DRAWING_RELS_PN = "drawings/_rels/drawing%d.xml.rels"

  # vml drawing part
  VML_DRAWING_PN = "drawings/vmlDrawing%d.vml"

  # drawing part
  TABLE_PN = "tables/table%d.xml"

  # pivot table parts
  PIVOT_TABLE_PN = "pivotTables/pivotTable%d.xml"

  # pivot table cache definition part name
  PIVOT_TABLE_CACHE_DEFINITION_PN = "pivotCache/pivotCacheDefinition%d.xml"

  # pivot table rels parts
  PIVOT_TABLE_RELS_PN = "pivotTables/_rels/pivotTable%d.xml.rels"

  # chart part
  CHART_PN = "charts/chart%d.xml"

  # chart part
  IMAGE_PN = "media/image%d.%s"

  # comment part
  COMMENT_PN = "comments%d.xml"

  # location of schema files for validation
  SCHEMA_BASE = File.dirname(__FILE__)+'/../../schema/'

  # App validation schema
  APP_XSD = SCHEMA_BASE + "shared-documentPropertiesExtended.xsd"

  # core validation schema
  CORE_XSD = SCHEMA_BASE + "opc-coreProperties.xsd"

  # content types validation schema
  CONTENT_TYPES_XSD = SCHEMA_BASE + "opc-contentTypes.xsd"

  # rels validation schema
  RELS_XSD = SCHEMA_BASE + "opc-relationships.xsd"

  # spreadsheetML validation schema
  SML_XSD = SCHEMA_BASE + "sml.xsd"

  # drawing validation schema
  DRAWING_XSD = SCHEMA_BASE + "dml-spreadsheetDrawing.xsd"

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
  ERR_RESTRICTION = "Invalid Data: %s. %s must be one of %s."

  # error message DataTypeValidator
  ERR_TYPE = "Invalid Data %s for %s. must be %s."

  # error message for RegexValidator
  ERR_REGEX = "Invalid Data. %s does not match %s."

  # error message for RangeValidator
  ERR_RANGE = "Invalid Data. %s must be between %s and %s, (inclusive:%s) you gave: %s"

  # error message for sheets that use a name which is longer than 31 bytes
  ERR_SHEET_NAME_TOO_LONG = "Your worksheet name '%s' is too long. Worksheet names must be 31 characters (bytes) or less"

  # error message for sheets that use a name which include invalid characters
  ERR_SHEET_NAME_CHARACTER_FORBIDDEN = "Your worksheet name '%s' contains a character which is not allowed by MS Excel and will cause repair warnings. Please change the name of your sheet."

  # error message for duplicate sheet names
  ERR_DUPLICATE_SHEET_NAME = "There is already a worksheet in this workbook named '%s'. Please use a unique name"

  # error message when user does not provide color and or style options for border in Style#add_sytle
  ERR_INVALID_BORDER_OPTIONS = "border hash must include both style and color. e.g. :border => { :color => 'FF000000', :style => :thin }. You provided: %s"

  # error message for invalid border id reference
  ERR_INVALID_BORDER_ID = "The border id you specified (%s) does not exist. Please add a border with Style#add_style before referencing its index."

  # error message for invalid angles
  ERR_ANGLE = "Angles must be a value between -90 and 90. You provided: %s"

  # error message for non 'integerish' value
  ERR_INTEGERISH = "You value must be, or be castable via to_i, an Integer. You provided %s"

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
  pattern = "[\x0-\x08\x0B\x0C\x0E-\x1F]"
  pattern= pattern.respond_to?(:encode) ? pattern.encode('UTF-8') : pattern

  # The regular expression used to remove control characters from worksheets
  CONTROL_CHAR_REGEX = Regexp.new(pattern, 'n')

end
