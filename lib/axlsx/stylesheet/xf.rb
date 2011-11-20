module Axlsx
  # The Xf class defines a formatting record for use in Styles
  class Xf
    #does not support extList (ExtensionList)

    # The cell alignment for this style
    # @return [CellAlignment]
    # @see CellAlignment
    attr_accessor :alignment

    # The cell protection for this style
    # @return [CellProtection]
    # @see CellProtection
    attr_accessor :protection

    # id of the numFmt to apply to this style
    # @return [Integer]
    attr_accessor :numFmtId

    # index (0 based) of the font to be used in this style
    # @return [Integer]
    attr_accessor :fontId
    
    # index (0 based) of the fill to be used in this style
    # @return [Integer]
    attr_accessor :fillId

    # index (0 based) of the border to be used in this style
    # @return [Integer]
    attr_accessor :borderId

    # index (0 based) of cellStylesXfs item to be used in this style. Only applies to cellXfs items
    # @return [Integer]
    attr_accessor :xfId

    # indecates if text should be prefixed by a single quote in the cell
    # @return [Boolean]
    attr_accessor :quotePrefix

    # indicates if the cell has a pivot table drop down button
    # @return [Boolean]
    attr_accessor :pivotButton

    # indicates if the numFmtId should be applied
    # @return [Boolean]
    attr_accessor :applyNumberFormat

    # indicates if the fontId should be applied
    # @return [Boolean]
    attr_accessor :applyFont
    
    # indicates if the fillId should be applied
    # @return [Boolean]
    attr_accessor :applyFill

    # indicates if the borderId should be applied
    # @return [Boolean]
    attr_accessor :applyBorder

    # Indicates if the alignment options should be applied
    # @return [Boolean]
    attr_accessor :applyAlignment

    # Indicates if the protection options should be applied
    # @return [Boolean]
    attr_accessor :applyProtection

    # Creates a new Xf object
    # @option [Integer] numFmtId
    # @option [Integer] fontId
    # @option [Integer] fillId
    # @option [Integer] borderId
    # @option [Integer] xfId
    # @option [Boolean] quotePrefix
    # @option [Boolean] pivotButton
    # @option [Boolean] applyNumberFormat 
    # @option [Boolean] applyFont
    # @option [Boolean] applyFill
    # @option [Boolean] applyBorder
    # @option [Boolean] applyAlignment
    # @option [Boolean] applyProtection
    # @option [CellAlignment] alignment
    # @option [CellProtection] protection
    def initialize(options={})
      options.each do |o|
        self.send("#{o[0]}=", o[1]) if self.respond_to? "#{o[0]}="
      end
    end    
    
    def alignment=(v) DataTypeValidator.validate "Xf.alignment", CellAlignment, v; @alignment = v end
    def protection=(v) DataTypeValidator.validate "Xf.protection", CellProtection, v; @protection = v end

    def numFmtId=(v) Axlsx::validate_unsigned_int v; @numFmtId = v end    
    def fontId=(v) Axlsx::validate_unsigned_int v; @fontId = v end    
    def fillId=(v) Axlsx::validate_unsigned_int v; @fillId = v end    
    def borderId=(v) Axlsx::validate_unsigned_int v; @borderId = v end    
    def xfId=(v) Axlsx::validate_unsigned_int v; @xfId = v end    
    def quotePrefix=(v) Axlsx::validate_boolean v; @quotePrefix = v end    
    def pivotButton=(v) Axlsx::validate_boolean v; @pivotButton = v end    
    def applyNumberFormat=(v) Axlsx::validate_boolean v; @applyNumberFormat = v end    
    def applyFont=(v) Axlsx::validate_boolean v; @applyFont = v end    
    def applyFill=(v) Axlsx::validate_boolean v; @applyFill = v end    
    def applyBorder=(v) Axlsx::validate_boolean v; @applyBorder = v end    
    def applyAlignment=(v) Axlsx::validate_boolean v; @applyAlignment = v end    
    def applyProtection=(v) Axlsx::validate_boolean v; @applyProtection = v end    

    # Serializes the xf elemen
    # @param [Nokogiri::XML::Builder] xml The document builder instance this objects xml will be added to.
    # @return [String]
    def to_xml(xml)
      xml.xf(self.instance_values.reject { |k, v| [:alignment, :protection, :extList, :name].include? k.to_sym}) {
        alignment.to_xml(xml) if self.alignment
        protection.to_xml(xml) if self.protection
      }
    end
  end
end
