# encoding: UTF-8
module Axlsx
  # The Xf class defines a formatting record for use in Styles. The recommended way to manage styles for your workbook is with Styles#add_style
  # @see Styles#add_style
  class Xf
    #does not support extList (ExtensionList)

    include Axlsx::SerializedAttributes
    include Axlsx::OptionsParser
    # Creates a new Xf object
    # @option options [Integer] numFmtId
    # @option options [Integer] fontId
    # @option options [Integer] fillId
    # @option options [Integer] borderId
    # @option options [Integer] xfId
    # @option options [Boolean] quotePrefix
    # @option options [Boolean] pivotButton
    # @option options [Boolean] applyNumberFormat
    # @option options [Boolean] applyFont
    # @option options [Boolean] applyFill
    # @option options [Boolean] applyBorder
    # @option options [Boolean] applyAlignment
    # @option options [Boolean] applyProtection
    # @option options [CellAlignment] alignment
    # @option options [CellProtection] protection
    def initialize(options={})
      parse_options options
    end

    serializable_attributes :numFmtId, :fontId, :fillId, :borderId, :xfId, :quotePrefix,
                            :pivotButton, :applyNumberFormat, :applyFont, :applyFill, :applyBorder, :applyAlignment,
                            :applyProtection

    # The cell alignment for this style
    # @return [CellAlignment]
    # @see CellAlignment
    attr_reader :alignment

    # The cell protection for this style
    # @return [CellProtection]
    # @see CellProtection
    attr_reader :protection

    # id of the numFmt to apply to this style
    # @return [Integer]
    attr_reader :numFmtId

    # index (0 based) of the font to be used in this style
    # @return [Integer]
    attr_reader :fontId

    # index (0 based) of the fill to be used in this style
    # @return [Integer]
    attr_reader :fillId

    # index (0 based) of the border to be used in this style
    # @return [Integer]
    attr_reader :borderId

    # index (0 based) of cellStylesXfs item to be used in this style. Only applies to cellXfs items
    # @return [Integer]
    attr_reader :xfId

    # indecates if text should be prefixed by a single quote in the cell
    # @return [Boolean]
    attr_reader :quotePrefix

    # indicates if the cell has a pivot table drop down button
    # @return [Boolean]
    attr_reader :pivotButton

    # indicates if the numFmtId should be applied
    # @return [Boolean]
    attr_reader :applyNumberFormat

    # indicates if the fontId should be applied
    # @return [Boolean]
    attr_reader :applyFont

    # indicates if the fillId should be applied
    # @return [Boolean]
    attr_reader :applyFill

    # indicates if the borderId should be applied
    # @return [Boolean]
    attr_reader :applyBorder

    # Indicates if the alignment options should be applied
    # @return [Boolean]
    attr_reader :applyAlignment

    # Indicates if the protection options should be applied
    # @return [Boolean]
    attr_reader :applyProtection

      # @see Xf#alignment
    def alignment=(v) DataTypeValidator.validate "Xf.alignment", CellAlignment, v; @alignment = v end

    # @see protection
    def protection=(v) DataTypeValidator.validate "Xf.protection", CellProtection, v; @protection = v end

    # @see numFmtId
    def numFmtId=(v) Axlsx::validate_unsigned_int v; @numFmtId = v end

    # @see fontId
    def fontId=(v) Axlsx::validate_unsigned_int v; @fontId = v end
    # @see fillId
    def fillId=(v) Axlsx::validate_unsigned_int v; @fillId = v end
    # @see borderId
    def borderId=(v) Axlsx::validate_unsigned_int v; @borderId = v end
    # @see xfId
    def xfId=(v) Axlsx::validate_unsigned_int v; @xfId = v end
    # @see quotePrefix
    def quotePrefix=(v) Axlsx::validate_boolean v; @quotePrefix = v end
    # @see pivotButton
    def pivotButton=(v) Axlsx::validate_boolean v; @pivotButton = v end
    # @see applyNumberFormat
    def applyNumberFormat=(v) Axlsx::validate_boolean v; @applyNumberFormat = v end
    # @see applyFont
    def applyFont=(v) Axlsx::validate_boolean v; @applyFont = v end
    # @see applyFill
    def applyFill=(v) Axlsx::validate_boolean v; @applyFill = v end

    # @see applyBorder
    def applyBorder=(v) Axlsx::validate_boolean v; @applyBorder = v end

    # @see applyAlignment
    def applyAlignment=(v) Axlsx::validate_boolean v; @applyAlignment = v end

    # @see applyProtection
    def applyProtection=(v) Axlsx::validate_boolean v; @applyProtection = v end

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      str << '<xf '
      serialized_attributes str
      str << '>'
      alignment.to_xml_string(str) if self.alignment
      protection.to_xml_string(str) if self.protection
      str << '</xf>'
    end


  end
end
