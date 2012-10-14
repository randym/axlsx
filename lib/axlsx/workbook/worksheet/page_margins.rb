module Axlsx
  # PageMargins specify the margins when printing a worksheet.
  #
  # For compatibility, PageMargins serialize to an empty string, unless at least one custom margin value
  # has been specified. Otherwise, it serializes to a PageMargin element specifying all 6 margin values
  # (using default values for margins that have not been specified explicitly).
  #
  # @note The recommended way to manage page margins is via Worksheet#page_margins
  # @see Worksheet#page_margins
  # @see Worksheet#initialize
  class PageMargins

    include Axlsx::OptionsParser
    include Axlsx::SerializedAttributes

    # Creates a new PageMargins object
    # @option options [Numeric] left The left margin in inches
    # @option options [Numeric] right The right margin in inches
    # @option options [Numeric] bottom The bottom margin in inches
    # @option options [Numeric] top The top margin in inches
    # @option options [Numeric] header The header margin in inches
    # @option options [Numeric] footer The footer margin in inches
    def initialize(options={})
      # Default values taken from MS Excel for Mac 2011
      @left = @right = DEFAULT_LEFT_RIGHT
      @top = @bottom = DEFAULT_TOP_BOTTOM
      @header = @footer = DEFAULT_HEADER_FOOTER
      parse_options options
    end

    serializable_attributes :left, :right, :bottom, :top, :header, :footer

    # Default left and right margin (in inches)
    DEFAULT_LEFT_RIGHT = 0.75

    # Default top and bottom margins (in inches)
    DEFAULT_TOP_BOTTOM = 1.00

    # Default header and footer margins (in inches)
    DEFAULT_HEADER_FOOTER = 0.50

    # Left margin (in inches)
    # @return [Float]
    attr_reader :left

    # Right margin (in inches)
    # @return [Float]
    attr_reader :right

    # Top margin (in inches)
    # @return [Float]
    attr_reader :top

    # Bottom margin (in inches)
    # @return [Float]
    attr_reader :bottom

    # Header margin (in inches)
    # @return [Float]
    attr_reader :header

    # Footer margin (in inches)
    # @return [Float]
    attr_reader :footer

    # Set some or all margins at once.
    # @param [Hash] margins the margins to set (possible keys are :left, :right, :top, :bottom, :header and :footer).
    def set(margins)
      margins.select do |k, v|
        next unless [:left, :right, :top, :bottom, :header, :footer].include? k
        send("#{k}=", v)
      end
    end

    # @see left
    def left=(v); Axlsx::validate_unsigned_numeric(v); @left = v end
    # @see right
    def right=(v); Axlsx::validate_unsigned_numeric(v); @right = v end
    # @see top
    def top=(v); Axlsx::validate_unsigned_numeric(v); @top = v end
    # @see bottom
    def bottom=(v); Axlsx::validate_unsigned_numeric(v); @bottom = v end
    # @see header
    def header=(v); Axlsx::validate_unsigned_numeric(v); @header = v end
    # @see footer
    def footer=(v); Axlsx::validate_unsigned_numeric(v); @footer = v end

    # Serializes the page margins element
    # @param [String] str
    # @return [String]
    # @note For compatibility, this is a noop unless custom margins have been specified.
    # @see #custom_margins_specified?
    def to_xml_string(str = '')
      str << '<pageMargins '
      serialized_attributes str
      str << '/>'
    end
  end
end
