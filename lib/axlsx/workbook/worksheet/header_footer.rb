module Axlsx
  # Header/Footer options for printing a worksheet. All settings are optional.
  #
  # @note The recommended way of managing header/footers is via Worksheet#header_footer
  # @see Worksheet#initialize
  class HeaderFooter

    include Axlsx::OptionsParser
    include Axlsx::SerializedAttributes
    include Axlsx::Accessors

    # Creates a new HeaderFooter object
    # @option options [String] odd_header The content for headers on odd numbered pages.
    # @option options [String] odd_footer The content for footers on odd numbered pages.
    # @option options [String] even_header The content for headers on even numbered pages.
    # @option options [String] even_footer The content for footers on even numbered pages.
    # @option options [String] first_header The content for headers on even numbered pages.
    # @option options [String] first_footer The content for footers on even numbered pages.
    # @option options [Boolean] different_odd_even Setting this to true will show different headers/footers on odd and even pages. When false, the odd headers/footers are used on each page. (Default: false)
    # @option options [Boolean] different_first If true, will use the first header/footer on page 1. Otherwise, the odd header/footer is used.
    def initialize(options = {})
      parse_options options
    end

    serializable_attributes :different_odd_even, :different_first

    string_attr_accessor :odd_header, :odd_footer, :even_header, :even_footer, :first_header, :first_footer
    boolean_attr_accessor :different_odd_even, :different_first

    # Set some or all header/footers at once.
    # @param [Hash] options The header/footer options to set (possible keys are :odd_header, :odd_footer, :even_header, :even_footer, :first_header, :first_footer, :different_odd_even, and :different_first).
    def set(options)
      parse_options options
    end

    # Serializes the header/footer object.
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      str << "<headerFooter "
      serialized_attributes str
      str << ">"

      str << "<oddHeader>#{odd_header}</oddHeader>" unless odd_header.nil?
      str << "<oddFooter>#{odd_footer}</oddFooter>" unless odd_footer.nil?

      str << "<evenHeader>#{even_header}</evenHeader>" unless even_header.nil?
      str << "<evenFooter>#{even_footer}</evenFooter>" unless even_footer.nil?

      str << "<firstHeader>#{first_header}</firstHeader>" unless first_header.nil?
      str << "<firstFooter>#{first_footer}</firstFooter>" unless first_footer.nil?

      str << "</headerFooter>"
    end
  end
end
