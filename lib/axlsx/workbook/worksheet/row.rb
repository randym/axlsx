# encoding: UTF-8
module Axlsx
  # A Row is a single row in a worksheet.
  # @note The recommended way to manage rows and cells is to use Worksheet#add_row
  # @see Worksheet#add_row
  class Row

    # The worksheet this row belongs to
    # @return [Worksheet]
    attr_reader :worksheet

    # The cells this row holds
    # @return [SimpleTypedList]
    attr_reader :cells

    # The height of this row in points, if set explicitly.
    # @return [Float]
    attr_reader :height

    # TODO  18.3.1.73
    # collapsed
    # customFormat
    # hidden
    # outlineLevel
    # ph
    # s (style)
    # spans
    # thickTop
    # thickBottom


    # Creates a new row. New Cell objects are created based on the values, types and style options.
    # A new cell is created for each item in the values array. style and types options are applied as follows:
    #   If the types option is defined and is a symbol it is applied to all the cells created.
    #   If the types option is an array, cell types are applied by index for each cell
    #   If the types option is not set, the cell will automatically determine its type.
    #   If the style option is defined and is an Integer, it is applied to all cells created.
    #   If the style option is an array, style is applied by index for each cell.
    #   If the style option is not defined, the default style (0) is applied to each cell.
    # @param [Worksheet] worksheet
    # @option options [Array] values
    # @option options [Array, Symbol] types
    # @option options [Array, Integer] style
    # @option options [Float] height the row's height (in points)
    # @see Row#array_to_cells
    # @see Cell
    def initialize(worksheet, values=[], options={})
      @height = nil
      self.worksheet = worksheet
      @cells = SimpleTypedList.new Cell
      @worksheet.rows << self
      self.height = options.delete(:height) if options[:height]
      array_to_cells(values, options)
    end

    # The index of this row in the worksheet
    # @return [Integer]
    def index
      worksheet.rows.index(self)
    end

    def to_xml_string
      if custom_height?
        '<row r="' << (index+1).to_s << '" customHeight="1" ht="' << height.to_s << '">' << @cells.map { |cell| cell.to_xml_string }.join << '</row>'
      else
        '<row r="' << (index+1).to_s << '">' << (@cells.map { |cell| cell.to_xml_string }.join) << '</row>'
      end
    end
    # Serializes the row
    # @param [Nokogiri::XML::Builder] xml The document builder instance this objects xml will be added to.
    # @return [String]
    def to_xml(xml)
      attrs = {:r => index+1}
      attrs.merge!(:customHeight => 1, :ht => height) if custom_height?
      xml.row(attrs) { |ixml| @cells.each { |cell| cell.to_xml(ixml) } }
    end

    # Adds a singel sell to the row based on the data provided and updates the worksheet's autofit data.
    # @return [Cell]
    def add_cell(value="", options={})
      c = Cell.new(self, value, options)
      update_auto_fit_data
      c
    end

    # sets the style for every cell in this row
    def style=(style)
      cells.each_with_index do | cell, index |
        s = style.is_a?(Array) ? style[index] : style
        cell.style = s
      end
    end

    # returns the cells in this row as an array
    # This lets us transpose the rows into columns
    # @return [Array]
    def to_ary
      @cells.to_ary
    end

    # @see height
    def height=(v); Axlsx::validate_unsigned_numeric(v) unless v.nil?; @height = v end

    # true if the row height has been manually set
    # @return [Boolean]
    # @see #height
    def custom_height?
      @height != nil
    end

    private

    # assigns the owning worksheet for this row
    def worksheet=(v) DataTypeValidator.validate "Row.worksheet", Worksheet, v; @worksheet=v; end

    # Tell the worksheet to update autofit data for the columns based on this row's cells.
    # @return [SimpleTypedList]
    def update_auto_fit_data
      worksheet.send(:update_auto_fit_data, self.cells)
    end


    # Converts values, types, and style options into cells and associates them with this row.
    # A new cell is created for each item in the values array.
    # If value option is defined and is a symbol it is applied to all the cells created.
    # If the value option is an array, cell types are applied by index for each cell
    # If the style option is defined and is an Integer, it is applied to all cells created.
    # If the style option is an array, style is applied by index for each cell.
    # @option options [Array] values
    # @option options [Array, Symbol] types
    # @option options [Array, Integer] style
    def array_to_cells(values, options={})
      values = values
      DataTypeValidator.validate 'Row.array_to_cells', Array, values
      types, style = options.delete(:types), options.delete(:style)
      values.each_with_index do |value, index|
        cell_style = style.is_a?(Array) ? style[index] : style
        options[:style] = cell_style if cell_style
        cell_type = types.is_a?(Array)? types[index] : types
        options[:type] = cell_type if cell_type
        Cell.new(self, value, options)
      end
    end
  end

end
