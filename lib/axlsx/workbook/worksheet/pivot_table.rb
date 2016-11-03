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
      @data_sheet = nil
      @rows = []
      @columns = []
      @data = []
      @pages = []
      # Sort by proc eventually used by pivot fields, per pivot field name
      # Hash<String,Proc>
      @pivot_field_sort_by_procs = {}
      @subtotal = nil
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

    # The sheet used as data source for the pivot table
    # @return [Worksheet]
    attr_writer :data_sheet

    # @see #data_sheet
    def data_sheet
      @data_sheet || @sheet
    end

    # The range where the data for this pivot table lives.
    # @return [String]
    attr_reader :range

    # (see #range)
    def range=(v)
      DataTypeValidator.validate "#{self.class}.range", [String], v
      if v.is_a?(String)
        @range = v
      end
    end

    # The rows
    # @return [Array]
    attr_reader :rows


    # (see #rows)
    def rows=(v)
      DataTypeValidator.validate "#{self.class}.rows", [Array], v
      v.each do |ref|
        DataTypeValidator.validate "#{self.class}.rows[]", [String], ref
      end
      @rows = v
    end

    # The columns
    # @return [Array]
    attr_reader :columns

    # (see #columns)
    def columns=(v)
      DataTypeValidator.validate "#{self.class}.columns", [Array], v
      v.each do |ref|
        DataTypeValidator.validate "#{self.class}.columns[]", [String], ref
      end
      @columns = v
    end

    # The data
    # @return [Array]
    attr_reader :data

    # (see #data)
    def data=(v)
      DataTypeValidator.validate "#{self.class}.data", [Array], v
      @data = []
      v.each do |data_field|
        if data_field.is_a? String
          data_field = {:ref => data_field}
        end
        data_field.values.each do |value|
          DataTypeValidator.validate "#{self.class}.data[]", [String], value
        end
        @data << data_field
      end
      @data
    end

    # The pages
    # @return [String]
    attr_reader :pages

    # (see #pages)
    def pages=(v)
      DataTypeValidator.validate "#{self.class}.pages", [Array], v
      v.each do |ref|
        DataTypeValidator.validate "#{self.class}.pages[]", [String], ref
      end
      @pages = v
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

    # The cache_definition for this pivot table
    # @return [PivotTableCacheDefinition]
    def cache_definition
      @cache_definition ||= PivotTableCacheDefinition.new(self)
    end

    # The relationships for this pivot table.
    # @return [Relationships]
    def relationships
      r = Relationships.new
      r << Relationship.new(cache_definition, PIVOT_TABLE_CACHE_DEFINITION_R, "../#{cache_definition.pn}")
      r
    end

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      str << '<?xml version="1.0" encoding="UTF-8"?>'
      str << ('<pivotTableDefinition xmlns="' << XML_NS << '" name="' << name << '" cacheId="' << cache_definition.cache_id.to_s << '"  dataOnRows="0" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Data" showMultipleLabel="0" showMemberPropertyTips="0" useAutoFormatting="1" indent="0" compact="0" compactData="0" gridDropZones="1" multipleFieldFilters="0" mergeItem="1" fieldListSortAscending="1">')
      str << ('<location firstDataCol="1" firstDataRow="1" firstHeaderRow="1" ref="' << ref << '"/>')
      str << ('<pivotFields count="' << header_cells_count.to_s << '">')
      header_cell_values.each do |cell_value|
        str << pivot_field_for(cell_value)
      end
      str << '</pivotFields>'
      offset = data.size > 1 ? 1 : 0
      if rows.empty?
        str << '<rowItems count="1"><i/></rowItems>'
      else
        str << ('<rowFields count="' << rows.size.to_s << '">')
        rows.each do |row_value|
          str << ('<field x="' << header_index_of(row_value).to_s << '"/>')
        end
        str << '</rowFields>'
        str << "<rowItems count=\"#{rows.size}\">"
        rows.size.times { str << '<i/>' }
        str << '</rowItems>'
      end
      if columns.empty?
        str << '<colFields count="1"><field x="-2"/></colFields>' if data.size > 1
      else
        str << ('<colFields count="' << (columns.size + 1).to_s << '">')
        columns.each do |column_value|
          str << ('<field x="' << header_index_of(column_value).to_s << '"/>')
        end
        str << ('<field x="-2"/>') if data.size > 1
        str << '</colFields>'
        str << "<colItems count=\"#{columns.size}\">"
        columns.size.times { str << '<i/>' }
        str << '</colItems>'
      end
      unless pages.empty?
        str << ('<pageFields count="' << pages.size.to_s << '">')
        pages.each do |page_value|
          str << ('<pageField fld="' << header_index_of(page_value).to_s << '"/>')
        end
        str << '</pageFields>'
      end
      unless data.empty?
        str << "<dataFields count=\"#{data.size}\">"
        data.each do |datum_value|
          str << "<dataField name='Sum of #{datum_value[:ref]}' fld='#{header_index_of(datum_value[:ref])}' baseField='0' baseItem='0'"
          str << " subtotal='#{datum_value[:subtotal]}' " if datum_value[:subtotal]
          str << "/>"
        end
        str << '</dataFields>'
      end
      str << '<pivotTableStyleInfo name="PivotStyleMedium11" showLastColumn="1" showColStripes="0" showRowStripes="0" showColHeaders="1" showRowHeaders="1"/>'
      str << '</pivotTableDefinition>'
    end

    # References for header cells
    # @return [Array]
    def header_cell_refs
      Axlsx::range_to_a(header_range).first
    end

    # The header cells for the pivot table
    # @return [Array]
    def header_cells
      data_sheet[header_range]
    end

    # The values in the header cells collection
    # @return [Array]
    def header_cell_values
      header_cells.map(&:value)
    end

    # The number of cells in the header_cells collection
    # @return [Integer]
    def header_cells_count
      header_cells.count
    end

    # The index of a given value in the header cells
    # @return [Integer]
    def header_index_of(value)
      header_cell_values.index(value)
    end

    # Return unique values for a given header name
    # @return [Array]
    def unique_values_for_header(header_name)
      header_idx = header_index_of(header_name)
      headers_count = header_cells_count
      data_sheet[range, in_bound: true].map(&:value)[headers_count..-1].each_slice(headers_count).map { |row| row[header_idx] }.uniq
    end

    # Set an order by proc for a given pivot field
    def pivot_field_sort_by(field_name, &block)
      @pivot_field_sort_by_procs[field_name] = block
    end

    private

    def pivot_field_for(cell_ref)
      if rows.include? cell_ref
        axis_pivot_field_for('axisRow', cell_ref)
      elsif columns.include? cell_ref
        axis_pivot_field_for('axisCol', cell_ref)
      elsif pages.include? cell_ref
        axis_pivot_field_for('axisCol', cell_ref)
      elsif data_refs.include? cell_ref
        '<pivotField dataField="1" compact="0" outline="0" subtotalTop="0" showAll="0" includeNewItemsInFilter="1">' + '</pivotField>'
      else
        '<pivotField compact="0" outline="0" subtotalTop="0" showAll="0" includeNewItemsInFilter="1">' + '</pivotField>'
      end
    end

    # Return tag to be used for a pivot field along a given axis
    # @return [String]
    def axis_pivot_field_for(axis_name, cell_ref)
      unique_values = unique_values_for_header(cell_ref)
      indexes = if @pivot_field_sort_by_procs.key?(cell_ref)
        # Sort the values of the pivot field
        unique_values.each_with_index.sort_by { |value, idx| @pivot_field_sort_by_procs[cell_ref].call(value) }.map(&:last)
      else
        unique_values.size.times.to_a
      end
      "<pivotField axis=\"#{axis_name}\" compact=\"0\" outline=\"0\" subtotalTop=\"0\" showAll=\"0\" includeNewItemsInFilter=\"1\">" +
        "<items count=\"#{unique_values.size+1}\">" +
        indexes.map { |idx| "<item x=\"#{idx}\"/>" }.join +
        '<item t="default"/>' +
        '</items>' +
        '</pivotField>'
    end

    def data_refs
      data.map { |hash| hash[:ref] }
    end

    def header_range
      range.gsub(/^(\w+?)(\d+)\:(\w+?)\d+$/, '\1\2:\3\2')
    end
  end
end
