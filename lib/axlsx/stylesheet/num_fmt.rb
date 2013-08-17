# encoding: UTF-8
module Axlsx
  # A NumFmt object defines an identifier and formatting code for data in cells.
  # @note The recommended way to manage styles is Styles#add_style
  class NumFmt

    include Axlsx::OptionsParser
    include Axlsx::SerializedAttributes

    # Creates a new NumFmt object
    # @param [Hash] options Options for the number format object
    # @option [Integer] numFmtId The predefined format id or new format id for this format
    # @option [String] formatCode The format code for this number format
    def initialize(options={})
      @numFmtId = 0
      @formatCode = ""
      parse_options options
    end

    serializable_attributes :formatCode, :numFmtId

    # @return [String] The formatting to use for this number format.
    # @see http://support.microsoft.com/kb/264372
    attr_reader :formatCode

    # @return [Integer] An unsigned integer referencing a standard or custom number format.
    # @note
    #  These are the known formats I can dig up. The constant NUM_FMT_PERCENT is 9, and uses the default % formatting. Axlsx also defines a few formats for date and time that are commonly used in asia as NUM_FMT_YYYYMMDD and NUM_FRM_YYYYMMDDHHMMSS.
    #   1 0
    #   2 0.00
    #   3 #,##0
    #   4 #,##0.00
    #   5 $#,##0_);($#,##0)
    #   6 $#,##0_);[Red]($#,##0)
    #   7 $#,##0.00_);($#,##0.00)
    #   8 $#,##0.00_);[Red]($#,##0.00)
    #   9 0%
    #   10 0.00%
    #   11 0.00E+00
    #   12 # ?/?
    #   13 # ??/??
    #   14 m/d/yyyy
    #   15 d-mmm-yy
    #   16 d-mmm
    #   17 mmm-yy
    #   18 h:mm AM/PM
    #   19 h:mm:ss AM/PM
    #   20 h:mm
    #   21 h:mm:ss
    #   22 m/d/yyyy h:mm
    #   37 #,##0_);(#,##0)
    #   38 #,##0_);[Red](#,##0)
    #   39 #,##0.00_);(#,##0.00)
    #   40 #,##0.00_);[Red](#,##0.00)
    #   45 mm:ss
    #   46 [h]:mm:ss
    #   47 mm:ss.0
    #   48 ##0.0E+0
    #   49 @
    # @see Axlsx
    attr_reader :numFmtId

    # @see numFmtId
    def numFmtId=(v) Axlsx::validate_unsigned_int v; @numFmtId = v end

    # @see formatCode
    def formatCode=(v) Axlsx::validate_string v; @formatCode = v end

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      str << '<numFmt '
      serialized_attributes str
      str << '/>'
    end

  end
end
