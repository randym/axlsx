# encoding: UTF-8
module Axlsx

  # The Col class defines column attributes for columns in sheets.
  class Col

    # Maximum column width limit in MS Excel is 255 characters
    # https://support.microsoft.com/en-us/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3
    MAX_WIDTH = 255

    include Axlsx::OptionsParser
    include Axlsx::SerializedAttributes
    # Create a new Col objects
    # @param min First column affected by this 'column info' record.
    # @param max Last column affected by this 'column info' record.
    # @option options [Boolean] collapsed see Col#collapsed
    # @option options [Boolean] hidden see Col#hidden
    # @option options [Boolean] outlineLevel see Col#outlineLevel
    # @option options [Boolean] phonetic see Col#phonetic
    # @option options [Integer] style see Col#style
    # @option options [Numeric] width see Col#width
    def initialize(min, max, options={})
      Axlsx.validate_unsigned_int(max)
      Axlsx.validate_unsigned_int(min)
      @min = min
      @max = max
      parse_options options
    end

    serializable_attributes :collapsed, :hidden, :outline_level, :phonetic, :style, :width, :min, :max, :best_fit, :custom_width

    # First column affected by this 'column info' record.
    # @return [Integer]
    attr_reader :min

    # Last column affected by this 'column info' record.
    # @return [Integer]
    attr_reader :max

    # Flag indicating if the specified column(s) is set to 'best fit'. 'Best fit' is set to true under these conditions:
    # The column width has never been manually set by the user, AND The column width is not the default width
    # 'Best fit' means that when numbers are typed into a cell contained in a 'best fit' column, the column width should
    #  automatically resize to display the number. [Note: In best fit cases, column width must not be made smaller, only larger. end note]
    # @return [Boolean]
    attr_reader :best_fit
    alias :bestFit :best_fit

    # Flag indicating if the outlining of the affected column(s) is in the collapsed state.
    # @return [Boolean]
    attr_reader :collapsed

    # Flag indicating if the affected column(s) are hidden on this worksheet.
    # @return [Boolean]
    attr_reader :hidden

    # Outline level of affected column(s). Range is 0 to 7.
    # @return [Integer]
    attr_reader :outline_level
    alias :outlineLevel :outline_level

    # Flag indicating if the phonetic information should be displayed by default for the affected column(s) of the worksheet.
    # @return [Boolean]
    attr_reader :phonetic

    # Default style for the affected column(s). Affects cells not yet allocated in the column(s). In other words, this style applies to new columns.
    # @return [Integer]
    attr_reader :style

    # The width of the column
    # @return [Numeric]
    attr_reader :width

    # @return [Boolean]
    attr_reader :custom_width
    alias :customWidth :custom_width

    # @see Col#collapsed
    def collapsed=(v)
      Axlsx.validate_boolean(v)
      @collapsed = v
    end

    # @see Col#hidden
    def hidden=(v)
      Axlsx.validate_boolean(v)
      @hidden = v
    end

    # @see Col#outline
    def outline_level=(v)
      Axlsx.validate_unsigned_numeric(v)
      raise ArgumentError, 'outlineLevel must be between 0 and 7' unless 0 <= v && v <= 7
      @outline_level = v
    end
    alias :outlineLevel= :outline_level=

    # @see Col#phonetic
    def phonetic=(v)
      Axlsx.validate_boolean(v)
      @phonetic = v
    end

    # @see Col#style
    def style=(v)
      Axlsx.validate_unsigned_int(v)
      @style = v
    end

   # @see Col#width
    def width=(v)
      # Removing this validation make a 10% difference in performance
      # as it is called EVERY TIME A CELL IS ADDED - the proper solution
      # is to only set this if a calculated value is greated than the
      # current @width value.
      # TODO!!!
      #Axlsx.validate_unsigned_numeric(v) unless v == nil
      @custom_width = @best_fit = v != nil
      @width = v.nil? ? v : [v, MAX_WIDTH].min
    end

    # updates the width for this col based on the cells autowidth and
    # an optionally specified fixed width
    # @param [Cell] cell The cell to use in updating this col's width
    # @param [Integer] fixed_width If this is specified the width is set
    # to this value and the cell's attributes are ignored.
    # @param [Boolean] use_autowidth If this is false, the cell's
    # autowidth value will be ignored.
    def update_width(cell, fixed_width=nil, use_autowidth=true)
      if fixed_width.is_a? Numeric
       self.width = fixed_width
      elsif use_autowidth
       cell_width = cell.autowidth
       self.width = cell_width unless (width || 0) > (cell_width || 0)
      end
    end

    # Serialize this columns data to an xml string
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      serialized_tag('col', str)
    end

  end
end
