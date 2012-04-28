# encoding: UTF-8
module Axlsx
  # Table
  # @note Worksheet#add_table is the recommended way to create tables for your worksheets.
  # @see README for examples
  class Table


    # The reference to the table data
    # @return [String]
    attr_reader :ref

    # The name of the table.
    # @return [String]
    attr_reader :name

    # The style for the table.
    # @return [TableStyle]
    attr_reader :style

    # Creates a new Table object
    # @param [String] ref The reference to the table data.
    # @param [Sheet] ref The sheet containing the table data.
    # @option options [Cell, String] name
    # @option options [TableStyle] style
    def initialize(ref, sheet, options={})
      @ref = ref
      @sheet = sheet
      @style = nil
      @sheet.workbook.tables << self
      @name = "Table#{index+1}"
      options.each do |o|
        self.send("#{o[0]}=", o[1]) if self.respond_to? "#{o[0]}="
      end
      yield self if block_given?
    end

    # The index of this chart in the workbooks charts collection
    # @return [Integer]
    def index
      @sheet.workbook.tables.index(self)
    end

    # The part name for this table
    # @return [String]
    def pn
      "#{TABLE_PN % (index+1)}"
    end

    # The relation reference id for this table
    # @return [String]
    def rId
      "rId#{index+1}"
    end

    # The name of the Table.
    # @param [String, Cell] v
    # @return [Title]
    def name=(v)
      DataTypeValidator.validate "#{self.class}.name", [String], v
      if v.is_a?(String)
        @name = v
      end
    end

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      str << '<?xml version="1.0" encoding="UTF-8"?>'
      str << '<table xmlns="' << XML_NS << '" id="' << (index+1).to_s << '" name="' << @name << '" displayName="' << @name.gsub(/\s/,'_') << '" '
      str << 'ref="' << @ref << '" totalsRowShown="0">'
      str << '<autoFilter ref="' << @ref << '"/>'
      str << '<tableColumns count="' << header_cells.length.to_s << '">'
      header_cells.each_with_index do |cell,index|
        str << '<tableColumn id ="' << (index+1).to_s << '" name="' << cell.value << '"/>'
      end
      str << '</tableColumns>'
      #TODO implement tableStyleInfo
      str << '<tableStyleInfo showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0" name="TableStyleMedium9" />'
      str << '</table>'
    end

    # The style for the table.
    # TODO
    # def style=(v) DataTypeValidator.validate "Table.style", Integer, v, lambda { |arg| arg >= 1 && arg <= 48 }; @style = v; end

    private

    # get the header cells (hackish)
    def header_cells
      header = @ref.gsub(/^(\w+?)(\d+)\:(\w+?)\d+$/, '\1\2:\3\2')
      @sheet[header]
    end
  end
end
