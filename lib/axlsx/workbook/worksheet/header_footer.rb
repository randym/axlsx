module Axlsx
  # Header and footer strings. All settings are optional.
  #
  # For a page indicator, supply "Page &amp;P of &amp;N"
  #
  # @note The recommended way to manage print options is via Worksheet#header_footer
  # @see Worksheet#header_footer
  # @see Worksheet#initialize
  class HeaderFooter

    # Creates a new HeaderFooter object
    # @option options [String] left_header Header for left side of printed page
    # @option options [String] center_header Header for center of printed page
    # @option options [String] right_header Header for right side of printed page
    # @option options [String] left_footer Footer for left side of printed page
    # @option options [String] center_footer Footer for center of printed page
    # @option options [String] right_footer Footer for right side of printed page
    def initialize(options = {})
      set(options)
    end

    # Set some or all page settings at once.
    # @param [Hash] options The page settings to set (possible keys are :fit_to_height, :fit_to_width, :orientation, :paper_height, :paper_width, and :scale).
    def set(options)
      options.each do |k, v|
        send("#{k}=", v) if respond_to? "#{k}="
      end
    end

    # Header for left side of page.
    # @return [String]
    attr_reader :left_header

    # Header for right side of page.
    # @return [String]
    attr_reader :right_header

    # Header for center of page.
    # @return [String]
    attr_reader :center_header

    # Footer for left side of page.
    # @return [String]
    attr_reader :left_footer

    # Footer for right side of page.
    # @return [String]
    attr_reader :right_footer

    # Footer for center of page.
    # @return [String]
    attr_reader :center_footer

    # @see left_header
    def left_header=(v); @left_header = v; end
    # @see center_header
    def center_header=(v); @center_header = v; end
    # @see right_header
    def right_header=(v); @right_header = v; end

    # @see left_footer
    def left_footer=(v); @left_footer = v; end
    # @see center_footer
    def center_footer=(v); @center_footer = v; end
    # @see right_footer
    def right_footer=(v); @right_footer = v; end

    # Serializes the headerFooter block
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      str << '<headerFooter>'
      ['even','odd'].each do |which|
        str << "<#{which}Header>"
        if @left_header
          str << '&amp;L'
          str << @left_header
        end
        if @center_header
          str << '&amp;C'
          str << @center_header
        end
        if @right_header
          str << '&amp;R'
          str << @right_header
        end
        str << "</#{which}Header>"
        str << "<#{which}Footer>"
        if @left_footer
          str << '&amp;L'
          str << @left_footer
        end
        if @center_footer
          str << '&amp;C'
          str << @center_footer
        end
        if @right_footer
          str << "&amp;R"
          str << @right_footer
        end
        str << "</#{which}Footer>"
      end
      str << '</headerFooter>'
    end
  end
end
