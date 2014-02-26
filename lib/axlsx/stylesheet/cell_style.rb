# encoding: UTF-8
module Axlsx
  # CellStyle defines named styles that reference defined formatting records and can be used in your worksheet.
  # @note Using Styles#add_style is the recommended way to manage cell styling.
  # @see Styles#add_style
  class CellStyle

    include Axlsx::OptionsParser
    include Axlsx::SerializedAttributes

    # Creats a new CellStyle object
    # @option options [String] name
    # @option options [Integer] xfId
    # @option options [Integer] buildinId
    # @option options [Integer] iLevel
    # @option options [Boolean] hidden
    # @option options [Boolean] customBuiltIn
    def initialize(options={})
      parse_options options
    end

    serializable_attributes :name, :xfId, :buildinId, :iLevel, :hidden, :customBuilin

    # The name of this cell style
    # @return [String]
    attr_reader :name

    # The formatting record id this named style utilizes
    # @return [Integer]
    # @see Axlsx::Xf
    attr_reader :xfId

    # The buildinId to use when this named style is applied
    # @return [Integer]
    # @see Axlsx::NumFmt
    attr_reader :builtinId

    # Determines if this formatting is for an outline style, and what level of the outline it is to be applied to.
    # @return [Integer]
    attr_reader :iLevel

    # Determines if this named style should show in the list of styles when using excel
    # @return [Boolean]
    attr_reader :hidden

    # Indicates that the build in style reference has been customized.
    # @return [Boolean]
    attr_reader :customBuiltin

     # @see name
    def name=(v)  Axlsx::validate_string v; @name = v end
    # @see xfId
    def xfId=(v) Axlsx::validate_unsigned_int v; @xfId = v end
    # @see builtinId
    def builtinId=(v) Axlsx::validate_unsigned_int v; @builtinId = v end
    # @see iLivel
    def iLevel=(v) Axlsx::validate_unsigned_int v; @iLevel = v end
    # @see hidden
    def hidden=(v) Axlsx::validate_boolean v; @hidden = v end
    # @see customBuiltin
    def customBuiltin=(v) Axlsx::validate_boolean v; @customBuiltin = v end

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      serialized_tag('cellStyle', str)
    end

  end

end
