# -*- coding: utf-8 -*-
module Axlsx
  # A cell in a worksheet. 
  # Cell stores inforamation requried to serialize a single worksheet cell to xml. You must provde the Row that the cell belongs to and the cells value. The data type will automatically be determed if you do not specify the :type option. The default style will be applied if you do not supply the :style option. Changing the cell's type will recast the value to the type specified. Altering the cell's value via the property accessor will also automatically cast the provided value to the cell's type.
  # @example Manually creating and manipulating Cell objects
  #   ws = Workbook.new.add_worksheet 
  #   # This is the simple, and recommended way to create cells. Data types will automatically be determined for you.
  #   ws.add_row :values => [1,"fish",Time.now]
  #
  #   # but you can also do this
  #   r = ws.add_row
  #   r.add_cell 1
  # 
  #   # or even this
  #   r = ws.add_row
  #   c = Cell.new row, 1, :value=>integer
  #
  #   # cells can also be accessed via Row#cells. The example here changes the cells type, which will automatically updated the value from 1 to 1.0
  #   r.cells.last.type = :float
  # 
  # @note The recommended way to generate cells is via Worksheet#add_row
  # 
  # @see Worksheet#add_row
  class Cell

    # The index of the cellXfs item to be applied to this cell.
    # @return [Integer] 
    # @see Axlsx::Styles
    attr_reader :style

    # The row this cell belongs to.
    # @return [Row]
    attr_reader :row
    
    # The cell's data type. Currently only four types are supported, :time, :float, :integer and :string.
    # Changing the type for a cell will recast the value into that type. If no type option is specified in the constructor, the type is 
    # automatically determed.
    # @see Cell#cell_type_from_value
    # @return [Symbol] The type of data this cell's value is cast to. 
    # @raise [ArgumentExeption] Cell.type must be one of [:time, :float, :integer, :string]
    # @note 
    #  If the value provided cannot be cast into the type specified, type is changed to :string and the following logic is applied.
    #   :string to :integer or :float, type coversions always return 0 or 0.0    
    #   :string, :integer, or :float to :time conversions always return the original value as a string and set the cells type to :string.
    #  No support is currently implemented for parsing time strings.
    attr_reader :type
    # @see type
    def type=(v) 
      RestrictionValidator.validate "Cell.type", [:time, :float, :integer, :string], v      
      @type=v 
      self.value = @value
    end


    # The value of this cell.
    # @return casted value based on cell's type attribute.
    attr_reader :value
    # @see value
    def value=(v)
      #TODO: consider doing value based type determination first?
      @value = cast_value(v)
    end


    # @param [Row] row The row this cell belongs to.
    # @param [Any] value The value associated with this cell. 
    # @option options [Symbol] type The intended data type for this cell. If not specified the data type will be determined internally based on the vlue provided.
    # @option options [Integer] style The index of the cellXfs item to be applied to this cell. If not specified, the default style (0) will be applied.
    def initialize(row, value="", options={})
      self.row=row
      #reference for validation
      @styles = row.worksheet.workbook.styles
      @type= options[:type] || cell_type_from_value(value)
      self.style = options[:style] || 0 
      @value = cast_value(value)
      @row.cells << self
    end

    # @return [Integer] The index of the cell in the containing row.
    def index
      @row.cells.index(self)
    end

    # @return [String] The alpha(column)numeric(row) reference for this sell.
    # @example Relative Cell Reference
    #   ws.rows.first.cells.first.r #=> "A1" 
    def r
      "#{col_ref}#{@row.index+1}"      
    end

    # @return [String] The absolute alpha(column)numeric(row) reference for this sell.
    # @example Absolute Cell Reference
    #   ws.rows.first.cells.first.r #=> "$A$1" 
    def r_abs
      "$#{r.split('').join('$')}"
    end

    # @return [Integer] The cellXfs item index applied to this cell.
    # @raise [ArgumentError] Invalid cellXfs id if the value provided is not within cellXfs items range.
    def style=(v)
      Axlsx::validate_unsigned_int(v)
      count = @styles.cellXfs.size
      raise ArgumentError, "Invalid cellXfs id" unless v < count
      @style = v
    end



    # Serializes the cell
    # @param [Nokogiri::XML::Builder] xml The document builder instance this objects xml will be added to.
    # @return [String] xml text for the cell
    # @note
    #   Shared Strings are not used in this library. All values are set directly in the each sheet.
    def to_xml(xml)
      # Both 1.8 and 1.9 return the same 'fast_xf'
      # &#12491;&#12507;&#12531;&#12468;
      # &#12491;&#12507;&#12531;&#12468;
      
      # however nokogiri does a nice 'force_encoding' which we shall remove!
      if @type == :string 
        xml.c(:r => r, :t=>:inlineStr, :s=>style) { xml.is { xml.t @value.to_s } }
      else
        xml.c(:r => r, :s => style) { xml.v value }
      end
    end


    private 

    # assigns the owning row for this cell.
    def row=(v) DataTypeValidator.validate "Cell.row", Row, v; @row=v end
    
    # converts the column index into alphabetical values.
    # @note This follows the standard spreadsheet convention of naming columns A to Z, followed by AA to AZ etc.
    # @return [String]
    def col_ref
      chars = []
      index = self.index
      while index >= 26 do
        chars << ((index % 26) + 65).chr
        index /= 26
      end
      chars << ((chars.empty? ? index : index-1) + 65).chr
      chars.reverse.join
    end

    # Determines the cell type based on the cell value. 
    # @note This is only used when a cell is created but no :type option is specified, the following rules apply:
    #   1. If the value is an instance of Time, the type is set to :time
    #   2. :float and :integer types are determined by regular expression matching.
    #   3. Anything that does not meet either of the above is determined to be :string.
    # @return [Symbol] The determined type
    def cell_type_from_value(v)      
      if v.is_a? Time
        :time
      elsif v.to_s.match(/\A[+-]?\d+?\Z/) #numeric
        :integer
      elsif v.to_s.match(/\A[+-]?\d+\.\d+?\Z/) #float
        :float
      else
        :string
      end
    end

    # Cast the value into this cells data type. 
    # @note 
    #   About Time - Time in OOXML is *different* from what you might expect. The history as to why is interesting,  but you can safely assume that if you are generating docs on a mac, you will want to specify Workbook.1904 as true when using time typed values.
    # @see Axlsx#date1904
    def cast_value(v)
      if (@type == :time && v.is_a?(Time)) || (@type == :time && v.respond_to?(:to_time))
        v = v.respond_to?(:to_time) ? v.to_time : v
        epoc1900 = Time.local(1900,1,1)
        epoc1904 = Time.local(1904,1,1) 
        epoc = Workbook.date1904 ? epoc1904 : epoc1900
        ((v.localtime - epoc) /60.0/60.0/24.0).to_f
      elsif @type == :float
        v.to_f
      elsif @type == :integer
        v.to_i
      else
        @type = :string
        v.to_s
      end
    end    
  end
end
