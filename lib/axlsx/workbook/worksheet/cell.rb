# encoding: UTF-8
require 'cgi'
module Axlsx
  # A cell in a worksheet.
  # Cell stores inforamation requried to serialize a single worksheet cell to xml. You must provde the Row that the cell belongs to and the cells value. The data type will automatically be determed if you do not specify the :type option. The default style will be applied if you do not supply the :style option. Changing the cell's type will recast the value to the type specified. Altering the cell's value via the property accessor will also automatically cast the provided value to the cell's type.
  # @note The recommended way to generate cells is via Worksheet#add_row
  #
  # @see Worksheet#add_row
  class Cell

    include Axlsx::OptionsParser

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
    # @option options [Number] formula_value The value to cache for a formula cell.
    # @option options [Symbol] scheme must be one of :none, major, :minor
    def initialize(row, value = nil, options = {})
      @row = row
      # Do not use instance vars if not needed to use less RAM
      # And do not call parse_options on frequently used options
      # to get less GC cycles
      type = options.delete(:type) || cell_type_from_value(value)
      self.type = type unless type == :string


      val = options.delete(:style)
      self.style = val unless val.nil? || val == 0
      val = options.delete(:formula_value)
      self.formula_value = val unless val.nil?

      parse_options(options)

      self.value = value
      value.cell = self if contains_rich_text?
    end

    # this is the cached value for formula cells. If you want the values to render in iOS/Mac OSX preview
    # you need to set this.
    attr_accessor :formula_value

    # An array of available inline styes.
    # TODO change this to a hash where each key defines attr name and validator (and any info the validator requires)
    # then move it out to a module so we can re-use in in other classes.
    # needs to define bla=(v) and bla methods on the class that hook into a
    # set_attr method that kicks the suplied validator and updates the instance_variable
    # for the key
    INLINE_STYLES = [:value, :type, :font_name, :charset,
                     :family, :b, :i, :strike, :outline,
                     :shadow, :condense, :extend, :u,
                     :vertAlign, :sz, :color, :scheme].freeze

    # An array of valid cell types
    CELL_TYPES = [:date, :time, :float, :integer, :richtext,
                  :string, :boolean, :iso_8601, :text].freeze

    # The index of the cellXfs item to be applied to this cell.
    # @return [Integer]
    # @see Axlsx::Styles
    def style
      defined?(@style) ? @style : 0
    end

    # The row this cell belongs to.
    # @return [Row]
    attr_reader :row

    # The cell's data type.
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
    def type
      defined?(@type) ? @type : :string
    end

    # @see type
    def type=(v)
      RestrictionValidator.validate :cell_type, CELL_TYPES, v
      @type = v
      self.value = @value unless !defined?(@value) || @value.nil?
    end

    # The value of this cell.
    # @return [String, Integer, Float, Time, Boolean] casted value based on cell's type attribute.
    attr_reader :value

    # @see value
    def value=(v)
      #TODO: consider doing value based type determination first?
      @value = cast_value(v)
    end

    # Indicates that the cell has one or more of the custom cell styles applied.
    # @return [Boolean]
    def is_text_run?
      defined?(@is_text_run) && @is_text_run && !contains_rich_text?
    end

    def contains_rich_text?
      type == :richtext
    end

    # Indicates if the cell is good for shared string table
    def plain_string?
      (type == :string || type == :text) &&         # String typed
        !is_text_run? &&          # No inline styles
        !@value.nil? &&           # Not nil
        !@value.empty? &&         # Not empty
        !@value.start_with?(?=)  # Not a formula
    end

    # The inline font_name property for the cell
    # @return [String]
    attr_reader :font_name
    # @see font_name
    def font_name=(v) set_run_style :validate_string, :font_name, v; end

    # The inline charset property for the cell
    # As far as I can tell, this is pretty much ignored. However, based on the spec it should be one of the following:
    # 0 ￼ ANSI_CHARSET
    # 1 DEFAULT_CHARSET
    # 2 SYMBOL_CHARSET
    # 77 MAC_CHARSET
    # 128 SHIFTJIS_CHARSET
    # 129 ￼ HANGUL_CHARSET
    # 130 ￼ JOHAB_CHARSET
    # 134 ￼ GB2312_CHARSET
    # 136 ￼ CHINESEBIG5_CHARSET
    # 161 ￼ GREEK_CHARSET
    # 162 ￼ TURKISH_CHARSET
    # 163 ￼ VIETNAMESE_CHARSET
    # 177 ￼ HEBREW_CHARSET
    # 178 ￼ ARABIC_CHARSET
    # 186 ￼ BALTIC_CHARSET
    # 204 ￼ RUSSIAN_CHARSET
    # 222 ￼ THAI_CHARSET
    # 238 ￼ EASTEUROPE_CHARSET
    # 255 ￼ OEM_CHARSET
    # @return [String]
    attr_reader :charset
    # @see charset
    def charset=(v) set_run_style :validate_unsigned_int, :charset, v; end

    # The inline family property for the cell
    # @return [Integer]
    # 1 Roman
    # 2 Swiss
    # 3 Modern
    # 4 Script
    # 5 Decorative
    attr_reader :family
    # @see family
    def family=(v)
      set_run_style :validate_family, :family, v.to_i
    end

    # The inline bold property for the cell
    # @return [Boolean]
    attr_reader :b
    # @see b
    def b=(v) set_run_style :validate_boolean, :b, v; end

    # The inline italic property for the cell
    # @return [Boolean]
    attr_reader :i
    # @see i
    def i=(v) set_run_style :validate_boolean, :i, v; end

    # The inline strike property for the cell
    # @return [Boolean]
    attr_reader :strike
    # @see strike
    def strike=(v) set_run_style :validate_boolean, :strike, v; end

    # The inline outline property for the cell
    # @return [Boolean]
    attr_reader :outline
    # @see outline
    def outline=(v) set_run_style :validate_boolean, :outline, v; end

    # The inline shadow property for the cell
    # @return [Boolean]
    attr_reader :shadow
    # @see shadow
    def shadow=(v) set_run_style :validate_boolean, :shadow, v; end

    # The inline condense property for the cell
    # @return [Boolean]
    attr_reader :condense
    # @see condense
    def condense=(v) set_run_style :validate_boolean, :condense, v; end

    # The inline extend property for the cell
    # @return [Boolean]
    attr_reader :extend
    # @see extend
    def extend=(v) set_run_style :validate_boolean, :extend, v; end

    # The inline underline property for the cell.
    # It must be one of :none, :single, :double, :singleAccounting, :doubleAccounting, true
    # @return [Boolean]
    # @return [String]
    # @note true is for backwards compatability and is reassigned to :single
    attr_reader :u
    # @see u
    def u=(v)
      v = :single if (v == true || v == 1 || v == :true || v == 'true')
      set_run_style :validate_cell_u, :u, v
    end

    # The inline color property for the cell
    # @return [Color]
    attr_reader :color
    # @param [String] v The 8 character representation for an rgb color #FFFFFFFF"
    def color=(v)
      @color = v.is_a?(Color) ? v : Color.new(:rgb=>v)
      @is_text_run = true
    end

    # The inline sz property for the cell
    # @return [Inteter]
    attr_reader :sz
    # @see sz
    def sz=(v) set_run_style :validate_unsigned_int, :sz, v; end

    # The inline vertical alignment property for the cell
    # this must be one of [:baseline, :subscript, :superscript]
    # @return [Symbol]
    attr_reader :vertAlign
    # @see vertAlign
    def vertAlign=(v)
      RestrictionValidator.validate :cell_vertAlign, [:baseline, :subscript, :superscript], v
      set_run_style nil, :vertAlign, v
    end

    # The inline scheme property for the cell
    # this must be one of [:none, major, minor]
    # @return [Symbol]
    attr_reader :scheme
    # @see scheme
    def scheme=(v)
      RestrictionValidator.validate :cell_scheme, [:none, :major, :minor], v
      set_run_style nil, :scheme, v
    end

    # The Shared Strings Table index for this cell
    # @return [Integer]
    attr_reader :ssti

    # @return [Integer] The index of the cell in the containing row.
    def index
      @row.index(self)
    end

    # @return [String] The alpha(column)numeric(row) reference for this sell.
    # @example Relative Cell Reference
    #   ws.rows.first.cells.first.r #=> "A1"
    def r
      Axlsx::cell_r index, @row.row_index
    end

    # @return [String] The absolute alpha(column)numeric(row) reference for this sell.
    # @example Absolute Cell Reference
    #   ws.rows.first.cells.first.r #=> "$A$1"
    def r_abs
      "$#{r.match(%r{([A-Z]+)([0-9]+)})[1,2].join('$')}"
    end

    # @return [Integer] The cellXfs item index applied to this cell.
    # @raise [ArgumentError] Invalid cellXfs id if the value provided is not within cellXfs items range.
    def style=(v)
      Axlsx::validate_unsigned_int(v)
      count = styles.cellXfs.size
      raise ArgumentError, "Invalid cellXfs id" unless v < count
      @style = v
    end

    # @return [Array] of x/y coordinates in the sheet for this cell.
    def pos
      [index, row.row_index]
    end

    # Merges all the cells in a range created between this cell and the cell or string name for a cell  provided
    # @see worksheet.merge_cells
    # @param [Cell, String] target The last cell, or str ref for the cell in the merge range
    def merge(target)
      start, stop = if target.is_a?(String)
                      [self.r, target]
                    elsif(target.is_a?(Cell))
                      Axlsx.sort_cells([self, target]).map { |c| c.r }
                    end
      self.row.worksheet.merge_cells "#{start}:#{stop}" unless stop.nil?
    end

    # Serializes the cell
    # @param [Integer] r_index The row index for the cell
    # @param [Integer] c_index The cell index in the row.
    # @param [String] str The string index the cell content will be appended to. Defaults to empty string.
    # @return [String] xml text for the cell
    def to_xml_string(r_index, c_index, str = '')
      CellSerializer.to_xml_string r_index, c_index, self, str
    end

    def is_formula?
      type == :string && @value.to_s.start_with?(?=)
    end

    def is_array_formula?
      type == :string && @value.to_s.start_with?('{=') && @value.to_s.end_with?('}')
    end

    # returns the absolute or relative string style reference for
    # this cell.
    # @param [Boolean] absolute -when false a relative reference will be
    # returned.
    # @return [String]
    def reference(absolute=true)
      absolute ? r_abs : r
    end

    # Creates a defined name in the workbook for this cell.
    def name=(label)
      row.worksheet.workbook.add_defined_name "#{row.worksheet.name}!#{r_abs}", name: label
      @name = label
    end

    # returns the name of the cell
    attr_reader :name

    # Attempts to determine the correct width for this cell's content
    # @return [Float]
    def autowidth
      return if is_formula? || value.nil?
      if contains_rich_text?
        string_width('', font_size) + value.autowidth
      elsif styles.cellXfs[style].alignment && styles.cellXfs[style].alignment.wrap_text
        max_width = 0
        value.to_s.split(/\r?\n/).each do |line|
          width = string_width(line, font_size)
          max_width = width if width > max_width
        end
        max_width
      else
        string_width(value, font_size)
      end
    end

    # Returns the sanatized value
    # TODO find a better way to do this as it accounts for 30% of
    # processing time in benchmarking...
    def clean_value
      if (type == :string || type == :text) && !Axlsx::trust_input
        Axlsx::sanitize(::CGI.escapeHTML(@value.to_s))
      else
        @value.to_s
      end
    end

    private

    def styles
      row.worksheet.styles
    end

    # Returns the width of a string according to the current style
    # This is still not perfect...
    #  - scaling is not linear as font sizes increase
    def string_width(string, font_size)
      font_scale = font_size / 10.0
      (string.to_s.count(Worksheet::THIN_CHARS) + 3.0) * font_scale
    end

    # we scale the font size if bold style is applied to either the style font or
    # the cell itself. Yes, it is a bit of a hack, but it is much better than using
    # imagemagick and loading metrics for every character.
    def font_size
      return sz if sz
      font = styles.fonts[styles.cellXfs[style].fontId] || styles.fonts[0]
      (font.b || (defined?(@b) && @b)) ? (font.sz * 1.5) : font.sz
    end

    # Utility method for setting inline style attributes
    def set_run_style(validator, attr, value)
      return unless INLINE_STYLES.include?(attr.to_sym)
      Axlsx.send(validator, value) unless validator.nil?
      self.instance_variable_set :"@#{attr.to_s}", value
      @is_text_run = true
    end

    # @see ssti
    def ssti=(v)
      Axlsx::validate_unsigned_int(v)
      @ssti = v
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
      elsif v.to_s =~ Axlsx::NUMERIC_REGEX
        :integer
      elsif v.to_s =~ Axlsx::FLOAT_REGEX
        :float
      elsif v.to_s =~ Axlsx::ISO_8601_REGEX
        :iso_8601
      elsif v.is_a? RichText
        :richtext
      else
        :string
      end
    end

    # Cast the value into this cells data type.
    # @note
    #   About Time - Time in OOXML is *different* from what you might expect. The history as to why is interesting, but you can safely assume that if you are generating docs on a mac, you will want to specify Workbook.1904 as true when using time typed values.
    # @see Axlsx#date1904
    def cast_value(v)
      return v if v.is_a?(RichText) || v.nil?
      case type
      when :date
        self.style = STYLE_DATE if self.style == 0
        v
      when :time
        self.style = STYLE_DATE if self.style == 0
        if !v.is_a?(Time) && v.respond_to?(:to_time)
          v.to_time
        else
          v
        end
      when :float
        v.to_f
      when :integer
        v.to_i
      when :boolean
        v ? 1 : 0
      when :iso_8601
        #consumer is responsible for ensuring the iso_8601 format when specifying this type
        v
      else
        v.to_s
      end
    end

  end
end
