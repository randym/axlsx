# encoding: UTF-8
module Axlsx
  # A Row is a single row in a worksheet.
  # @note The recommended way to manage rows and cells is to use Worksheet#add_row
  # @see Worksheet#add_row
  class Row

    include SerializedAttributes
    include Accessors
    # No support is provided for the following attributes
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
      @ht = nil
      self.worksheet = worksheet
      @cells = SimpleTypedList.new Cell
      @worksheet.rows << self
      self.height = options.delete(:height) if options[:height]
      array_to_cells(values, options)
    end

    # A list of serializable attributes.
    serializable_attributes :hidden, :outline_level, :collapsed, :custom_format, :s, :ph, :custom_height, :ht

    # Boolean row attribute accessors
    boolean_attr_accessor :hidden, :collapsed, :custom_format, :ph, :custom_height

    # The worksheet this row belongs to
    # @return [Worksheet]
    attr_reader :worksheet

    # The cells this row holds
    # @return [SimpleTypedList]
    attr_reader :cells

    # Row height measured in point size. There is no margin padding on row height.
    # @return [Float]
    def height
      @ht
    end

    # Outlining level of the row, when outlining is on
    # @return [Integer]
    attr_reader :outline_level
    alias :outlineLevel :outline_level

    # The style applied ot the row. This affects the entire row.
    # @return [Integer]
    attr_reader :s

    # @see Row#s
    def s=(v)
      Axlsx.validate_unsigned_numeric(v)
      @custom_format = true 
      @s = v
    end

    # @see Row#outline
    def outline_level=(v)
      Axlsx.validate_unsigned_numeric(v)
      @outline_level = v
    end
    alias :outlineLevel= :outline_level=

    # The index of this row in the worksheet
    # @return [Integer]
    def index
      worksheet.rows.index(self)
    end

    # Serializes the row
    # @param [Integer] r_index The row index, 0 based.
    # @param [String] str The string this rows xml will be appended to.
    # @return [String]
    def to_xml_string(r_index, str = '')
      str << '<row '
      serialized_attributes(str, { :r => r_index + 1 })
      str << '>'
      @cells.each_with_index { |cell, c_index| cell.to_xml_string(r_index, c_index, str) }
      str << '</row>'
    end

    # Adds a singel sell to the row based on the data provided and updates the worksheet's autofit data.
    # @return [Cell]
    def add_cell(value="", options={})
      c = Cell.new(self, value, options)
      worksheet.send(:update_column_info, self.cells, [])
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
    def height=(v)
      Axlsx::validate_unsigned_numeric(v)
      unless v.nil?
        @ht = v
        @custom_height = true
      end
      @ht
    end

    private

    # assigns the owning worksheet for this row
    def worksheet=(v) DataTypeValidator.validate "Row.worksheet", Worksheet, v; @worksheet=v; end

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
      types, style, formula_values = options.delete(:types), options.delete(:style), options.delete(:formula_values)
      values.each_with_index do |value, index|

        #WTF IS THIS PAP?
        cell_style = style.is_a?(Array) ? style[index] : style
        options[:style] = cell_style if cell_style

        cell_type = types.is_a?(Array)? types[index] : types
        options[:type] = cell_type if cell_type

        formula_value = formula_values[index] if formula_values.is_a?(Array)
        options[:formula_value] = formula_value if formula_value

        Cell.new(self, value, options)

        options.delete(:style)
        options.delete(:type)
        options.delete(:formula_value)
      end
    end
  end

end
