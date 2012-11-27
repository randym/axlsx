# encoding: UTF-8
module Axlsx
  # Table
  # @note Worksheet#add_pivot_table is the recommended way to create tables for your worksheets.
  # @see README for examples
  class PivotTableCacheDefinition

    include Axlsx::OptionsParser

    # Creates a new PivotTable object
    # @param [String] pivot_table The pivot table this cache definition is in
    def initialize(pivot_table)
      @pivot_table = pivot_table
    end

    # # The reference to the pivot table data
    # # @return [PivotTable]
    attr_reader :pivot_table

    # The index of this chart in the workbooks charts collection
    # @return [Integer]
    def index
      pivot_table.sheet.workbook.pivot_tables.index(pivot_table)
    end

    # The part name for this table
    # @return [String]
    def pn
      "#{PIVOT_TABLE_CACHE_DEFINITION_PN % (index+1)}"
    end

    def cache_id
      index
    end

    # The relation reference id for this table
    # @return [String]
    def rId
      "rId#{index + 1}"
    end

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      str << '<?xml version="1.0" encoding="UTF-8"?>'
      str << '<pivotCacheDefinition xmlns="' << XML_NS << '" xmlns:r="' << XML_NS_R << '" invalid="1" refreshOnLoad="1" recordCount="0">'
      str <<   '<cacheSource type="worksheet">'
      str <<     '<worksheetSource ref="' << pivot_table.range << '" sheet="Data Sheet"/>'
      str <<   '</cacheSource>'
      str <<   '<cacheFields count="' << pivot_table.header_cells_count.to_s << '">'
      pivot_table.header_cells_count.times do |i|
        str <<   '<cacheField name="placeholder_' << i.to_s << '" numFmtId="0">'
        str <<     '<sharedItems count="0">'
        str <<     '</sharedItems>'
        str <<   '</cacheField>'
      end
      str <<   '</cacheFields>'
      str << '</pivotCacheDefinition>'
    end

  end
end
