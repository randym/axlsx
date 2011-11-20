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

    # The index of this row in the worksheet
    # @return [Integer]
    attr_reader :index

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
    # @see Row#array_to_cells
    # @see Cell
    def initialize(worksheet, values=[], options={})
      self.worksheet = worksheet
      @cells = SimpleTypedList.new Cell
      @worksheet.rows << self
      array_to_cells(values, options)
    end

    def index 
      worksheet.rows.index(self)
    end
    
    # Serializes the row
    # @param [Nokogiri::XML::Builder] xml The document builder instance this objects xml will be added to.
    # @return [String]
    def to_xml(xml)
      xml.row(:r => index+1) { @cells.each { |cell| cell.to_xml(xml) } }
    end

    # Adds a singel sell to the row based on the data provided and updates the worksheet's autofit data.
    # @return [Cell]
    def add_cell(value="", options={})
      c = Cell.new(self, value, options)
      update_auto_fit_data
      c
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
      types, style = options[:types], options[:style]
      values.each_with_index do |value, index|        
        cell_style = style.is_a?(Array) ? style[index] : style
        cell_type = types.is_a?(Array)? types[index] : types
        Cell.new(self, value, :style=>cell_style, :type=>cell_type) 
      end
    end
  end
  
end
