# encoding: UTF-8
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


    # An array of available inline styes.
    INLINE_STYLES = ['value', 'type', 'font_name', 'charset', 
                         'family', 'b', 'i', 'strike','outline', 
                         'shadow', 'condense', 'extend', 'u', 
                         'vertAlign', 'sz', 'color', 'scheme']


    # The index of the cellXfs item to be applied to this cell.
    # @return [Integer] 
    # @see Axlsx::Styles
    attr_reader :style

    # The row this cell belongs to.
    # @return [Row]
    attr_reader :row
    
    # The cell's data type. Currently only six types are supported, :date, :time, :float, :integer, :string and :boolean.
    # Changing the type for a cell will recast the value into that type. If no type option is specified in the constructor, the type is 
    # automatically determed.
    # @see Cell#cell_type_from_value
    # @return [Symbol] The type of data this cell's value is cast to. 
    # @raise [ArgumentExeption] Cell.type must be one of [:date, time, :float, :integer, :string, :boolean]
    # @note 
    #  If the value provided cannot be cast into the type specified, type is changed to :string and the following logic is applied.
    #   :string to :integer or :float, type conversions always return 0 or 0.0    
    #   :string, :integer, or :float to :time conversions always return the original value as a string and set the cells type to :string.
    #  No support is currently implemented for parsing time strings.
    attr_reader :type
    # @see type
    def type=(v) 
      RestrictionValidator.validate "Cell.type", [:date, :time, :float, :integer, :string, :boolean], v      
      @type=v 
      self.value = @value unless @value.nil?
    end


    # The value of this cell.
    # @return [String, Integer, Float, Time] casted value based on cell's type attribute.
    attr_reader :value
    # @see value
    def value=(v)
      #TODO: consider doing value based type determination first?
      @value = cast_value(v)
    end
    
    # The inline font_name property for the cell
    # @return [String]
    attr_reader :font_name
    # @see font_name
    def font_name=(v) Axlsx::validate_string(v); @font_name = v; end

    # The inline charset property for the cell
    # @return [String]
    attr_reader :charset
    # @see charset
    def charset=(v) Axlsx::validate_unsigned_int(v); @charset = v; end

    # The inline family property for the cell
    # @return [String]
    attr_reader :family
    # @see family
    def family=(v) Axlsx::validate_string(v); @family = v; end

    # The inline bold property for the cell
    # @return [Boolean]
    attr_reader :b
    # @see b
    def b=(v) Axlsx::validate_boolean(v); @b = v; end

    # The inline italic property for the cell
    # @return [Boolean]
    attr_reader :i
    # @see i
    def i=(v) Axlsx::validate_boolean(v); @i = v; end

    # The inline strike property for the cell
    # @return [Boolean]
    attr_reader :strike
    # @see strike
    def strike=(v) Axlsx::validate_boolean(v); @strike = v; end

    # The inline outline property for the cell
    # @return [Boolean]
    attr_reader :outline
    # @see outline
    def outline=(v) Axlsx::validate_boolean(v); @outline = v; end

    # The inline shadow property for the cell
    # @return [Boolean]
    attr_reader :shadow
    # @see shadow
    def shadow=(v) Axlsx::validate_boolean(v); @shadow = v; end

    # The inline condense property for the cell
    # @return [Boolean]
    attr_reader :condense
    # @see condense
    def condense=(v) Axlsx::validate_boolean(v); @condense = v; end

    # The inline extend property for the cell
    # @return [Boolean]
    attr_reader :extend
    # @see extend
    def extend=(v) Axlsx::validate_boolean(v); @extend = v; end

    # The inline underline property for the cell
    # @return [Boolean]
    attr_reader :u
    # @see u
    def u=(v) Axlsx::validate_boolean(v); @u = v; end

    # The inline color property for the cell
    # @return [Color]
    attr_reader :color
    # @param [String] The 8 character representation for an rgb color #FFFFFFFF"
    def color=(v) 
      @color = v.is_a?(Color) ? v : Color.new(:rgb=>v)
    end

    # The inline sz property for the cell
    # @return [Boolean]
    attr_reader :sz
    # @see sz
    def sz=(v) Axlsx::validate_unsigned_int(v); @sz = v; end

    # The inline vertical alignment property for the cell
    # this must be one of [:baseline, :subscript, :superscript]
    # @return [Symbol]
    attr_reader :vertAlign
    # @see vertAlign
    def vertAlign=(v) RestrictionValidator.validate "Cell.vertAlign", [:baseline, :subscript, :superscript], v; @vertAlign = v; end

    # The inline scheme property for the cell
    # this must be one of [:none, major, minor]
    # @return [Symbol]
    attr_reader :scheme
    # @see scheme
    def scheme=(v) RestrictionValidator.validate "Cell.schema", [:none, :major, :minor], v; @scheme = v; end

    # @param [Row] row The row this cell belongs to.
    # @param [Any] value The value associated with this cell. 
    # @option options [Symbol] type The intended data type for this cell. If not specified the data type will be determined internally based on the vlue provided.
    # @option options [Integer] style The index of the cellXfs item to be applied to this cell. If not specified, the default style (0) will be applied.
    # @option options [String] font_name
    # @option options [Integer] charset
    # @option options [String] family
    # @option options [Boolean] b
    # @option options [Boolean] i
    # @option options [Boolean] strike
    # @option options [Boolean] outline
    # @option options [Boolean] shadow
    # @option options [Boolean] condense
    # @option options [Boolean] extend
    # @option options [Boolean] u
    # @option options [Symbol] vertAlign must be one of :baseline, :subscript, :superscript
    # @option options [Integer] sz
    # @option options [String] color an 8 letter rgb specification
    # @option options [Symbol] scheme must be one of :none, major, :minor
    def initialize(row, value="", options={})
      self.row=row      
      @font_name = @charset = @family = @b = @i = @strike = @outline = @shadow = nil
      @condense = @u = @vertAlign = @sz = @color = @scheme = @extend = @ssti = nil
      @styles = row.worksheet.workbook.styles
      @row.cells << self      
      options.each do |o|
        self.send("#{o[0]}=", o[1]) if self.respond_to? "#{o[0]}="
      end
      @style ||= 0
      @type ||= cell_type_from_value(value)
      @value = cast_value(value)
    end

    # The Shared Strings Table index for this cell
    # @return [Integer]
    attr_reader :ssti
    
    # equality comparison to test value, type and inline style attributes
    # this is how we work out if the cell needs to be added or already exists in the shared strings table
    def shareable(v)

      #using reject becase 1.8.7 select returns an array...
      v_hash = v.instance_values.reject { |key, val| !INLINE_STYLES.include?(key) }
      self_hash = self.instance_values.reject { |key, val| !INLINE_STYLES.include?(key) }
      # required as color is an object, and the comparison will fail even though both use the same color.
      v_hash['color'] = v_hash['color'].instance_values if v_hash['color']
      self_hash['color'] = self_hash['color'].instance_values if self_hash['color']

      v_hash == self_hash
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

    # @return [Array] of x/y coordinates in the cheet for this cell.
    def pos
      [index, row.index]
    end

    # Merges all the cells in a range created between this cell and the cell or string name for a cell  provided
    # @see worksheet.merge_cells
    # @param [Cell, String] target The last cell, or str ref for the cell in the merge range
    def merge(target)
      range_end = if target.is_a?(String)
                    target
                  elsif(target.is_a?(Cell))
                    target.r
                  end
      self.row.worksheet.merge_cells "#{self.r}:#{range_end}" unless range_end.nil?
    end              

    # builds an xml text run based on this cells attributes. This is extracted from to_xml so that shared strings can use it.
    # @param [Nokogiri::XML::Builder] xml The document builder instance this output will be added to.
    # @return [String] the xml for this cell's text run
    def run_xml(xml)
      if (self.instance_values.keys & INLINE_STYLES).size > 0
        xml.r {
          xml.rPr {
            xml.rFont(:val=>@font_name) if @font_name
            xml.charset(:val=>@charset) if @charset
            xml.family(:val=>@family) if @family
            xml.b(:val=>@b) if @b
            xml.i(:val=>@i) if @i
            xml.strike(:val=>@strike) if @strike
            xml.outline(:val=>@outline) if @outline
            xml.shadow(:val=>@shadow) if @shadow
            xml.condense(:val=>@condense) if @condense
            xml.extend(:val=>@extend) if @extend
            @color.to_xml(xml) if @color
            xml.sz(:val=>@sz) if @sz
            xml.u(:val=>@u) if @u
            # :baseline, :subscript, :superscript
            xml.vertAlign(:val=>@vertAlign) if @vertAlign
            # :none, major, :minor
            xml.scheme(:val=>@scheme) if @scheme
          }
          xml.t @value.to_s
        }
      else
        xml.t @value.to_s
      end        
    end

    # Serializes the cell
    # @param [Nokogiri::XML::Builder] xml The document builder instance this objects xml will be added to.
    # @return [String] xml text for the cell
    def to_xml(xml)      
      if @type == :string 
        #parse formula
        if @value.start_with?('=')
          xml.c(:r => r, :t=>:str, :s=>style) {
            xml.f @value.to_s.gsub('=', '')
          }
        else
          #parse shared
          if @ssti
            xml.c(:r => r, :s=>style, :t => :s) { xml.v ssti }
          else
            #parse inline string
            xml.c(:r => r, :s=>style, :t => :inlineStr) {
              xml.is {
                run_xml(xml)
              }
            }
          end
        end
      elsif @type == :date
        # TODO: See if this is subject to the same restriction as Time below
        epoc = Workbook.date1904 ? Date.new(1904) : Date.new(1900)
        v = (@value-epoc).to_f
        xml.c(:r => r, :s => style) { xml.v v }
      elsif @type == :time
        # Using hardcoded offsets here as some operating systems will not except a 'negative' offset from the ruby epoc.
        epoc1900 = -2209021200 #Time.local(1900, 1, 1)
        epoc1904 = -2082877200 #Time.local(1904, 1, 1)
        epoc = Workbook.date1904 ? epoc1904 : epoc1900
        v = ((@value.localtime.to_f - epoc) /60.0/60.0/24.0).to_f
        xml.c(:r => r, :s => style) { xml.v v }
      elsif @type == :boolean
        xml.c(:r => r, :s => style, :t => :b) { xml.v value }
      else
        xml.c(:r => r, :s => style) { xml.v value }
      end
    end

    private 

    # @see ssti
    def ssti=(v) 
      Axlsx::validate_unsigned_int(v)
      @ssti = v
    end

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
    #   1. If the value is an instance of Date, the type is set to :date
    #   2. If the value is an instance of Time, the type is set to :time
    #   3. If the value is an instance of TrueClass or FalseClass, the type is set to :boolean
    #   4. :float and :integer types are determined by regular expression matching.
    #   5. Anything that does not meet either of the above is determined to be :string.
    # @return [Symbol] The determined type
    def cell_type_from_value(v)      
      if v.is_a?(Date)
        :date
      elsif v.is_a?(Time)
        :time
      elsif v.is_a?(TrueClass) || v.is_a?(FalseClass)
        :boolean
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
      if @type == :date
        self.style = STYLE_DATE if self.style == 0
        v
      elsif (@type == :time && v.is_a?(Time)) || (@type == :time && v.respond_to?(:to_time))
        self.style = STYLE_DATE if self.style == 0
        v.respond_to?(:to_time) ? v.to_time : v
      elsif @type == :float
        v.to_f
      elsif @type == :integer
        v.to_i
      elsif @type == :boolean
        v ? 1 : 0
      else
        @type = :string
        v.to_s
      end
    end    
  end
end
