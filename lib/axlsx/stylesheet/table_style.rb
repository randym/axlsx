# encoding: UTF-8
module Axlsx
  # A single table style definition and is a collection for tableStyleElements
  # @note Table are not supported in this version and only the defaults required for a valid workbook are created.
  class TableStyle < SimpleTypedList

    include Axlsx::OptionsParser
    include Axlsx::SerializedAttributes

    # creates a new TableStyle object
    # @raise [ArgumentError] if name option is not provided.
    # @param [String] name
    # @option options [Boolean] pivot
    # @option options [Boolean] table
    def initialize(name, options={})
      self.name = name
      parse_options options
      super TableStyleElement
    end

    serializable_attributes :name, :pivot, :table

    # The name of this table style
    # @return [string]
    attr_reader :name

    # indicates if this style should be applied to pivot tables
    # @return [Boolean]
    attr_reader :pivot

    # indicates if this style should be applied to tables
    # @return [Boolean]
    attr_reader :table

    # @see name
    def name=(v) Axlsx::validate_string v; @name=v end
    # @see pivot
    def pivot=(v) Axlsx::validate_boolean v; @pivot=v end
    # @see table
    def table=(v) Axlsx::validate_boolean v; @table=v end

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      str << '<tableStyle '
      serialized_attributes str, {:count => self.size}
      str << '>'
      each { |table_style_el| table_style_el.to_xml_string(str) }
      str << '</tableStyle>'
    end

  end
end
