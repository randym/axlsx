module Axlsx
  # Page setup settings for printing a worksheet. All settings are optional.
  #
  # @note The recommended way to manage print options is via Worksheet#page_setup
  # @see Worksheet#print_options
  # @see Worksheet#initialize
  class PageSetup

    # TODO: Attributes defined by Open XML spec that are not implemented yet:
    # 
    # * blackAndWhite
    # * cellComments
    # * copies
    # * draft
    # * errors
    # * firstPageNumber
    # * horizontalDpi
    # * pageOrder
    # * paperSize
    # * useFirstPageNumber
    # * usePrinterDefaults
    # * verticalDpi
    
    # Number of vertical pages to fit on.
    # @return [Integer]
    attr_reader :fit_to_height

    # Number of horizontal pages to fit on.
    # @return [Integer]
    attr_reader :fit_to_width

    # Orientation of the page (:default, :landscape, :portrait)
    # @return [Symbol]
    attr_reader :orientation 
    
    # Height of paper (string containing a number followed by a unit identifier: "297mm", "11in")
    # @return [String]
    attr_reader :paper_height

    # Width of paper (string containing a number followed by a unit identifier: "210mm", "8.5in")
    # @return [String]
    attr_reader :paper_width
    
    # Print scaling (percent value, given as integer ranging from 10 to 400)
    # @return [Integer]
    attr_reader :scale

    # The worksheet that owns this page_setup.
    # @return [Worksheet]
    attr_reader :worksheet

    # Creates a new PageSetup object
    # @option options [Worksheet] worksheet The worksheet that owns this page_setup **required**
    # @option options [Integer] fit_to_height Number of vertical pages to fit on
    # @option options [Integer] fit_to_width Number of horizontal pages to fit on
    # @option options [Symbol] orientation Orientation of the page (:default, :landscape, :portrait)
    # @option options [String] paper_height Height of paper (number followed by unit identifier: "297mm", "11in")
    # @option options [String] paper_width Width of paper (number followed by unit identifier: "210mm", "8.5in")
    # @option options [Integer] scale Print scaling (percent value, integer ranging from 10 to 400)
    def initialize(options = {})
      raise ArgumentError, "Worksheet option is required" unless options[:worksheet].is_a?(Worksheet)
      @worksheet = options[:worksheet]
      set(options)
    end

    # Set some or all page settings at once.
    # @param [Hash] options The page settings to set (possible keys are :fit_to_height, :fit_to_width, :orientation, :paper_height, :paper_width, and :scale).
    def set(options)
      options.each do |k, v|
        send("#{k}=", v) if respond_to? "#{k}="
      end
    end

    # @see fit_to_height
    def fit_to_height=(v); Axlsx::validate_unsigned_int(v); @fit_to_height = v; @worksheet.fit_to_page = true; end
    # @see fit_to_width
    def fit_to_width=(v); Axlsx::validate_unsigned_int(v); @fit_to_width = v; @worksheet.fit_to_page = true; end
    # @see orientation
    def orientation=(v); Axlsx::validate_page_orientation(v); @orientation = v; end
    # @see paper_height
    def paper_height=(v); Axlsx::validate_number_with_unit(v); @paper_height = v; end
    # @see paper_width
    def paper_width=(v); Axlsx::validate_number_with_unit(v); @paper_width = v; end
    # @see scale
    def scale=(v); Axlsx::validate_page_scale(v); @scale = v; end

    # Serializes the page settings element.
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      str << '<pageSetup '
      str << instance_values.reject{ |k, v| k == 'worksheet' }.map{ |k,v| k.gsub(/_(.)/){ $1.upcase } << %{="#{v}"} }.join(' ')
      str << '/>'
    end
  end
end
