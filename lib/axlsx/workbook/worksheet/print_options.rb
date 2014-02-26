module Axlsx
  # Options for printing a worksheet. All options are boolean and false by default.
  #
  # @note The recommended way to manage print options is via Worksheet#print_options
  # @see Worksheet#print_options
  # @see Worksheet#initialize
  class PrintOptions

    include Axlsx::OptionsParser
    include Axlsx::SerializedAttributes
    include Axlsx::Accessors
    # Creates a new PrintOptions object
    # @option options [Boolean] grid_lines Whether grid lines should be printed
    # @option options [Boolean] headings Whether row and column headings should be printed
    # @option options [Boolean] horizontal_centered Whether the content should be centered horizontally
    # @option options [Boolean] vertical_centered  Whether the content should be centered vertically
    def initialize(options = {})
      @grid_lines = @headings = @horizontal_centered = @vertical_centered = false
      set(options)
    end

    serializable_attributes :grid_lines, :headings, :horizontal_centered, :vertical_centered
    boolean_attr_accessor :grid_lines, :headings, :horizontal_centered, :vertical_centered

    # Set some or all options at once.
    # @param [Hash] options The options to set (possible keys are :grid_lines, :headings, :horizontal_centered, and :vertical_centered).
    def set(options)
      parse_options options
    end

    # Serializes the page options element.
    # @note As all attributes default to "false" according to the xml schema definition, the generated xml includes only those attributes that are set to true.
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      serialized_tag 'printOptions', str
    end
  end
end
