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

    # The identifier for this cache
    # @return [Integer]
    def cache_id
      index + 1
    end

    # The relationship id for this pivot table cache definition.
    # @see Relationship#Id
    # @return [String]
    def rId
      pivot_table.relationships.for(self).Id
    end

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      str << '<?xml version="1.0" encoding="UTF-8"?>'
      str << ('<pivotCacheDefinition xmlns="' << XML_NS << '" xmlns:r="' << XML_NS_R << '" invalid="1" refreshOnLoad="1" recordCount="0">')
      str <<   '<cacheSource type="worksheet">'
      str << (    '<worksheetSource ref="' << pivot_table.range << '" sheet="' << pivot_table.data_sheet.name << '"/>')
      str <<   '</cacheSource>'
      str << (  '<cacheFields count="' << pivot_table.header_cells_count.to_s << '">')
      pivot_table.header_cells.each do |cell|
        str << (  '<cacheField name="' << cell.value << '" numFmtId="0">')
        header_name = cell.value
        unique_values = (pivot_table.rows.include?(header_name) || pivot_table.columns.include?(header_name) ? pivot_table.unique_values_for_header(header_name) : [])
        contains_string = false
        contains_number = false
        contains_integer = false
        min_value = nil
        max_value = nil
        items_tags = []
        unique_values.each do |value|
          if value.is_a?(Numeric)
            tag_name = 'n'
            contains_number = true
            contains_integer = true if value.is_a?(Integer)
            min_value = value if min_value.nil? || value < min_value
            max_value = value if max_value.nil? || value > max_value
          else
            tag_name = 's'
            contains_string = true
          end
          items_tags << "<#{tag_name} v=\"#{value}\"/>"
        end
        str <<     "<sharedItems count=\"#{unique_values.size}\" containsSemiMixedTypes=\"#{contains_number ? '0' : '1'}\" containsInteger=\"#{contains_integer ? '1' : '0'}\" containsNumber=\"#{contains_number ? '1' : '0'}\" containsString=\"#{contains_string ? '1' : '0'}\"#{min_value.nil? ? '' : " minValue=\"#{min_value}\""}#{max_value.nil? ? '' : " maxValue=\"#{max_value}\""}>"
        str << items_tags.join
        str <<     '</sharedItems>'
        str <<   '</cacheField>'
      end
      str <<   '</cacheFields>'
      str << '</pivotCacheDefinition>'
    end

  end
end
