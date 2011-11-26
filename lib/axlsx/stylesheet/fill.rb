module Axlsx
  # The Fill is a formatting object that manages the background color, and pattern for cells.
  # @note The recommended way to manage styles in your workbook is to use Styles#add_style.
  # @see Styles#add_style
  # @see PatternFill
  # @see GradientFill
  class Fill

    # The type of fill
    # @return [PatternFill, GradientFill]
    attr_reader :fill_type

    # Creates a new Fill object
    # @param [PatternFill, GradientFill] fill_type 
    # @raise [ArgumentError] if the fill_type parameter is not a PatternFill or a GradientFill instance
    def initialize(fill_type)
      self.fill_type = fill_type
    end

    # Serializes the fill
    # @param [Nokogiri::XML::Builder] xml The document builder instance this objects xml will be added to.
    # @return [String]
    def to_xml(xml)
      xml.fill { @fill_type.to_xml(xml) }
    end

    # @see fill_type
    def fill_type=(v) DataTypeValidator.validate "Fill.fill_type", [PatternFill, GradientFill], v; @fill_type = v; end


  end
end
