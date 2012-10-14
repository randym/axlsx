# encoding: UTF-8
module Axlsx
  # TableStyles represents a collection of style definitions for table styles and pivot table styles.
  # @note Support for custom table styles does not exist in this version. Many of the classes required are defined in preparation for future release. Please do not attempt to add custom table styles.
  class TableStyles < SimpleTypedList

    include Axlsx::SerializedAttributes

    # Creates a new TableStyles object that is a container for TableStyle objects
    # @option options [String] defaultTableStyle
    # @option options [String] defaultPivotStyle
    def initialize(options={})
      @defaultTableStyle = options[:defaultTableStyle] || "TableStyleMedium9"
      @defaultPivotStyle = options[:defaultPivotStyle] || "PivotStyleLight16"
      super TableStyle
    end

    serializable_attributes :defaultTableStyle, :defaultPivotStyle

    # The default table style. The default value is 'TableStyleMedium9'
    # @return [String]
    attr_reader :defaultTableStyle

    # The default pivot table style. The default value is  'PivotStyleLight6'
    # @return [String]
    attr_reader :defaultPivotStyle

   # @see defaultTableStyle
    def defaultTableStyle=(v) Axlsx::validate_string(v); @defaultTableStyle = v; end
    # @see defaultPivotStyle
    def defaultPivotStyle=(v) Axlsx::validate_string(v); @defaultPivotStyle = v; end

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      str << '<tableStyles '
      serialized_attributes str, {:count => self.size }
      str << '>'
      each { |table_style| table_style.to_xml_string(str) }
      str << '</tableStyles>'
    end

  end

end
