module Axlsx
  # CellStyle defines named styles that reference defined formatting records and can be used in your worksheet.
  # @note Using Styles#add_style is the recommended way to manage cell styling.
  # @see Styles#add_style
  class CellStyle
    # The name of this cell style
    # @return [String]
    attr_accessor :name
    
    # The formatting record id this named style utilizes
    # @return [Integer]
    # @see Axlsx::Xf
    attr_accessor :xfId

    # The buildinId to use when this named style is applied
    # @return [Integer]
    # @see Axlsx::NumFmt
    attr_accessor :builtinId

    # Determines if this formatting is for an outline style, and what level of the outline it is to be applied to.
    # @return [Integer]
    attr_accessor :iLevel

    # Determines if this named style should show in the list of styles when using excel
    # @return [Boolean]
    attr_accessor :hidden

    # Indicates that the build in style reference has been customized.
    # @return [Boolean]
    attr_accessor :customBuiltin

    # Creats a new CellStyle object
    # @option options [String] name
    # @option options [Integer] xfId
    # @option options [Integer] buildinId
    # @option options [Integer] iLevel
    # @option options [Boolean] hidden
    # @option options [Boolean] customBuiltIn
    def initialize(options={})
      options.each do |o|
        self.send("#{o[0]}=", o[1]) if self.respond_to? o[0]
      end
    end

    def name=(v)  Axlsx::validate_string v; @name = v end
    def xfId=(v) Axlsx::validate_unsigned_int v; @xfId = v end
    def builtinId=(v) Axlsx::validate_unsigned_int v; @builtinId = v end
    def iLevel=(v) Axlsx::validate_unsigned_int v; @iLevel = v end
    def hidden=(v) Axlsx::validate_boolean v; @hidden = v end
    def customBuiltin=(v) Axlsx::validate_boolean v; @customBuiltin = v end

    # Serializes the cell style
    # @param [Nokogiri::XML::Builder] xml The document builder instance this objects xml will be added to.
    # @return [String]
    def to_xml(xml)
      xml.cellStyle(self.instance_values)
    end
  end

end
