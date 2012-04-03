# encoding: UTF-8
module Axlsx

  # The Col class defines column attributes for columns in sheets.
  class Col

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
    attr_reader :bestFit

    # Flag indicating if the outlining of the affected column(s) is in the collapsed state.
    # @return [Boolean]
    attr_reader :collapsed

    # Flag indicating if the affected column(s) are hidden on this worksheet.
    # @return [Boolean]
    attr_reader :hidden

    # Outline level of affected column(s). Range is 0 to 7.
    # @return [Integer]
    attr_reader :outlineLevel

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
    attr_reader :customWidth

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
    def outlineLevel=(v)
      Axlsx.validate_boolean(v)
      @outlineLevel = v
    end

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
      Axlsx.validate_unsigned_numeric(v) unless v == nil
      @customWidth = @bestFit = v != nil
      @width = v
    end

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
      options.each do |o|
        self.send("#{o[0]}=", o[1]) if self.respond_to? "#{o[0]}="
      end
    end

    # Serialize this columns data to an xml string
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      attrs = self.instance_values.reject{ |key, value| value == nil }
      str << '<col ' << attrs.map { |key, value| '' << key << '="' << value.to_s << '"' }.join(' ') << '/>'
    end

  end
end
