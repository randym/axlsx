module Axlsx

  #Sheet formatting properties
  # <xsd:complexType name="CT_SheetFormatPr">
  #   <xsd:attribute name="baseColWidth" type="xsd:unsignedInt" use="optional" default="8"/>
  #   <xsd:attribute name="defaultColWidth" type="xsd:double" use="optional"/>
  #   <xsd:attribute name="defaultRowHeight" type="xsd:double" use="required"/>
  #   <xsd:attribute name="customHeight" type="xsd:boolean" use="optional" default="false"/>
  #   <xsd:attribute name="zeroHeight" type="xsd:boolean" use="optional" default="false"/>
  #   <xsd:attribute name="thickTop" type="xsd:boolean" use="optional" default="false"/>
  #   <xsd:attribute name="thickBottom" type="xsd:boolean" use="optional" default="false"/>
  #   <xsd:attribute name="outlineLevelRow" type="xsd:unsignedByte" use="optional" default="0"/>
  #   <xsd:attribute name="outlineLevelCol" type="xsd:unsignedByte" use="optional" default="0"/>
  #</xsd:complexType>

  class SheetFormatPr
    include Axlsx::SerializedAttributes
    include Axlsx::OptionsParser
    include Axlsx::Accessors

    # creates a new sheet_format_pr object
    # @param [Hash] options initialization options
    # @option [Integer] base_col_width Specifies the number of characters of the maximum digit width of the normal style's font. This value does not include margin padding or extra padding for gridlines. It is only the number of characters.
    # @option [Float] default_col_width Default column width measured as the number of characters of the maximum digit width of the normal style's font.
    # @option [Float] default_row_height Default row height measured in point size. Optimization so we don't have to write the height on all rows. This can be written out if most rows have custom height, to achieve the optimization.
    # @option [Boolean] custom_height 'True' if defaultRowHeight value has been manually set, or is different from the default value.
    # @option [Boolean] zero_height 'True' if rows are hidden by default. This setting is an optimization used when most rows of the sheet are hidden.
    # @option [Boolean] think_top 'True' if rows have a thick top border by default.
    # @option [Boolean] thick_bottom 'True' if rows have a thick bottom border by default.
    # @option [Integer] outline_level_row Highest number of outline level for rows in this sheet. These values shall be in synch with the actual sheet outline levels.
    # @option [Integer] outline_level_col Highest number of outline levels for columns in this sheet. These values shall be in synch with the actual sheet outline levels.
    def initialize(options={})
      set_defaults
      parse_options options
    end

    serializable_attributes :base_col_width, :default_col_width, :default_row_height,
      :custom_height, :zero_height, :thick_top, :thick_bottom,
      :outline_level_row, :outline_level_col

    float_attr_accessor :default_col_width, :default_row_height

    boolean_attr_accessor :custom_height, :zero_height, :thick_top, :thick_bottom

    unsigned_int_attr_accessor :base_col_width, :outline_level_row, :outline_level_col 

    # serializes this object to an xml string
    # @param [String] str The string this objects serialization will be appended to
    # @return [String]
    def to_xml_string(str='')
      str << "<sheetFormatPr #{serialized_attributes}/>"
    end

    private
    def set_defaults
      @base_col_width = 8
      @default_row_height = 18
    end
  end
end
