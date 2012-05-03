module Axlsx
  # Options for printing a worksheet. All options are boolean and false by default.
  #
  # @note The recommended way to manage print options is via Worksheet#print_options
  # @see Worksheet#print_options
  # @see Worksheet#initialize
  class PrintOptions

    # Whether grid lines should be printed.
    # @return [Boolean]
    attr_reader :grid_lines

    # Whether row and column headings should be printed.
    # @return [Boolean]
    attr_reader :headings

    # Whether the content should be centered horizontally on the page.
    # @return [Boolean]
    attr_reader :horizontal_centered

    # Whether the content should be centered vertically on the page.
    # @return [Boolean]
    attr_reader :vertical_centered

    # Creates a new PrintOptions object
    # @option options [Boolean] grid_lines Whether grid lines should be printed
    # @option options [Boolean] headings Whether row and column headings should be printed
    # @option options [Boolean] horizontal_centered Whether the content should be centered horizontally
    # @option options [Boolean] vertical_centered  Whether the content should be centered vertically
    def initialize(options = {})
      @grid_lines = @headings = @horizontal_centered = @vertical_centered = false
      set(options)
    end

    # Set some or all options at once.
    # @param [Hash] options The options to set (possible keys are :grid_lines, :headings, :horizontal_centered, and :vertical_centered).
    def set(options)
      options.each do |k, v|
        send("#{k}=", v) if respond_to? "#{k}="
      end
    end

    # @see grid_lines
    def grid_lines=(v); Axlsx::validate_boolean(v); @grid_lines = v; end
    # @see headings
    def headings=(v); Axlsx::validate_boolean(v); @headings = v; end
    # @see horizontal_centered
    def horizontal_centered=(v); Axlsx::validate_boolean(v); @horizontal_centered = v; end
    # @see vertical_centered
    def vertical_centered=(v); Axlsx::validate_boolean(v); @vertical_centered = v; end

    # Serializes the page options element.
    # @note As all attributes default to "false" according to the xml schema definition, the generated xml includes only those attributes that are set to true.
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      str << '<printOptions '
      # 
      str << instance_values.select{ |k,v| v == true }.map{ |k,v| k.gsub(/_(.)/){ $1.upcase } << %{="#{v}"} }.join(' ')
      str << '/>'
    end
  end
end
