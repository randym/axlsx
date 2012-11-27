# encoding: UTF-8
module Axlsx
  # Table
  # @note Worksheet#add_pivot_table is the recommended way to create tables for your worksheets.
  # @see README for examples
  class PivotTable

    include Axlsx::OptionsParser

    # Creates a new PivotTable object
    # @param [String] ref The reference to where the pivot table lives like 'G4:L17'.
    # @param [String] range The reference to the pivot table data like 'A1:D31'.
    # @param [Worksheet] sheet The sheet containing the table data.
    # @option options [Cell, String] name
    # @option options [TableStyle] style
    def initialize(ref, range, sheet, options={})
      @ref = ref
      self.range = range
      @sheet = sheet
      @sheet.workbook.pivot_tables << self
      @name = "PivotTable#{index+1}"
      parse_options options
      yield self if block_given?
    end

    # The reference to the table data
    # @return [String]
    attr_reader :ref

    # The name of the table.
    # @return [String]
    attr_reader :name

    # The name of the sheet.
    # @return [String]
    attr_reader :sheet

    # The range where the data for this pivot table lives.
    # @return [String]
    attr_reader :range

    def range=(v)
      DataTypeValidator.validate "#{self.class}.range", [String], v
      if v.is_a?(String)
        @range = v
      end
    end

    # The index of this chart in the workbooks charts collection
    # @return [Integer]
    def index
      @sheet.workbook.pivot_tables.index(self)
    end

    # The part name for this table
    # @return [String]
    def pn
      "#{PIVOT_TABLE_PN % (index+1)}"
    end

    # The relationship part name of this pivot table
    # @return [String]
    def rels_pn
      "#{PIVOT_TABLE_RELS_PN % (index+1)}"
    end

    def header_cells_count
      header_cells.count
    end

    def cache_definition
      @cache_definition ||= PivotTableCacheDefinition.new(self)
    end

    # The worksheet relationships. This is managed automatically by the worksheet
    # @return [Relationships]
    def relationships
      r = Relationships.new
      r << Relationship.new(PIVOT_TABLE_CACHE_DEFINITION_R, "../#{cache_definition.pn}")
      r
    end

    # identifies the index of an object withing the collections used in generating relationships for the worksheet
    # @param [Any] object the object to search for
    # @return [Integer] The index of the object
    def relationships_index_of(object)
      objects = [cache_definition]
      objects.index(object)
    end

    # The relation reference id for this table
    # @return [String]
    def rId
      "rId#{index+1}"
    end

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      str << '<?xml version="1.0" encoding="UTF-8"?>'
      str << '<pivotTableDefinition xmlns="' << XML_NS << '" name="' << name << '" cacheId="' << cache_definition.cache_id.to_s << '"  dataOnRows="1" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Data" showMultipleLabel="0" showMemberPropertyTips="0" useAutoFormatting="1" indent="0" compact="0" compactData="0" gridDropZones="1" multipleFieldFilters="0">'
      str <<   '<location firstDataCol="1" firstDataRow="1" firstHeaderRow="1" ref="' << ref << '"/>'
      str <<   '<pivotFields count="' << header_cells_count.to_s << '">'
      header_cells_count.times do
        str <<   '<pivotField compact="0" outline="0" subtotalTop="0" showAll="0" includeNewItemsInFilter="1"/>'
      end
      str <<   '</pivotFields>'
      str << '</pivotTableDefinition>'
    end

    private

    # get the header cells (hackish)
    def header_cells
      header = range.gsub(/^(\w+?)(\d+)\:(\w+?)\d+$/, '\1\2:\3\2')
      @sheet[header]
    end

  end
end
