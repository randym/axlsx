# encoding: UTF-8
module Axlsx
  # A Row is a single row in a worksheet.
  # @note The recommended way to manage rows and cells is to use Worksheet#add_row
  # @see Worksheet#add_row
  class Row

    # No support is provided for the following attributes
    # spans
    # thickTop
    # thickBottom


    # A list of serilizable attributes.
    # @note height(ht) and customHeight are manages separately for now. Have a look at Row#height
    SERIALIZABLE_ATTRIBUTES = [:hidden, :outlineLevel, :collapsed, :s, :customFormat, :ph]
    
    # The worksheet this row belongs to
    # @return [Worksheet]
    attr_reader :worksheet

    # The cells this row holds
    # @return [SimpleTypedList]
    attr_reader :cells

    # Row height measured in point size. There is no margin padding on row height.
    # @return [Float]
    attr_reader :height

    # Flag indicating if the outlining of row.
    # @return [Boolean]
    attr_reader :collapsed

    # Flag indicating if the the row is hidden.
    # @return [Boolean]
    attr_reader :hidden

    # Outlining level of the row, when outlining is on
    # @return [Integer]
    attr_reader :outlineLevel

    # The style applied ot the row. This affects the entire row.
    # @return [Integer]
    attr_reader :s

    # indicates that a style has been applied directly to the row via Row#s
    # @return [Boolean]
    attr_reader :customFormat

    # indicates if the row should show phonetic
    # @return [Boolean]
    attr_reader :ph

    # NOTE removing this from the api as it is actually incorrect. 
    # having a method to style a row's cells is fine, but it is not an attribute on the row. 
    # The proper attribute is ':s'
    # attr_reader style
    #
    
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

    # @see Row#collapsed
    def collapsed=(v)
      Axlsx.validate_boolean(v)
      @collapsed = v
    end

    # @see Row#hidden
    def hidden=(v)
      Axlsx.validate_boolean(v)
      @hidden = v
    end
   
    # @see Row#ph
    def ph=(v) Axlsx.validate_boolean(v); @ph = v end

    # @see Row#s
    def s=(v) Axlsx.validate_unsigned_numeric(v); @s = v; @customFormat = true end

    # @see Row#outline
    def outlineLevel=(v)
      Axlsx.validate_unsigned_numeric(v)
      @outlineLevel = v
    end

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
      str << '<row r="' << (r_index + 1 ).to_s << '" '
      instance_values.select { |key, value| SERIALIZABLE_ATTRIBUTES.include? key.to_sym }.each do |key, value|
        str << key << '="' << value.to_s << '" '
      end
      if custom_height?
        str << 'customHeight="1" ht="' << height.to_s << '">'
      else
        str << '>'
      end
      @cells.each_with_index { |cell, c_index| cell.to_xml_string(r_index, c_index, str) }
      str << '</row>'
      str
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
        options.delete(:style)
        options.delete(:type)
      end
    end
  end

end
