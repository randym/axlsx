# encoding: UTF-8
module Axlsx
  # A single table style definition and is a collection for tableStyleElements
  # @note Table are not supported in this version and only the defaults required for a valid workbook are created.
  class TableStyle < SimpleTypedList

    # The name of this table style
    # @return [string]
    attr_reader :name

    # indicates if this style should be applied to pivot tables
    # @return [Boolean]
    attr_reader :pivot

    # indicates if this style should be applied to tables
    # @return [Boolean]
    attr_reader :table

    # creates a new TableStyle object
    # @raise [ArgumentError] if name option is not provided.
    # @param [String] name
    # @option options [Boolean] pivot
    # @option options [Boolean] table
    def initialize(name, options={})
      self.name = name
      options.each do |o|
        self.send("#{o[0]}=", o[1]) if self.respond_to? o[0]
      end
      super TableStyleElement
    end

    # @see name
    def name=(v) Axlsx::validate_string v; @name=v end
    # @see pivot
    def pivot=(v) Axlsx::validate_boolean v; @pivot=v end
    # @see table
    def table=(v) Axlsx::validate_boolean v; @table=v end

    # Serializes the table style
    # @param [Nokogiri::XML::Builder] xml The document builder instance this objects xml will be added to.
    # @return [String]
    def to_xml(xml)
      attr = self.instance_values.select { |k, v| [:name, :pivot, :table].include? k }
      attr[:count] = self.size
      xml.tableStyle(attr) { self.each { |table_style_el| table_style_el.to_xml(xml) } }
    end
  end
end
