module Axlsx
  # The color class represents a color used for borders, fills an fonts
  class Color
    # Determines if the color is system color dependant
    # @return [Boolean]
    attr_reader :auto
    
    # Backwards compatability color index
    # return [Integer]
    #attr_reader :indexed

    # The color as defined in rgb terms. 
    # @note 
    #  rgb colors need to conform to ST_UnsignedIntHex. That basically means put 'FF' before you color
    # @example rgb colors
    #  "FF000000" is black
    #  "FFFFFFFF" is white
    # @return [String]
    attr_reader :rgb

    # no support for theme just yet
    # @return [Integer]
    #attr_reader :theme
    
    # The tint value.
    # @note valid values are between -1.0 and 1.0
    # @return [Float]
    attr_reader :tint
    
    # Creates a new Color object
    # @option options [Boolean] auto
    # @option options [String] rgb
    # @option options [Float] tint    
    def initialize(options={})
      @rgb = "FF000000"
      options.each do |o|
        self.send("#{o[0]}=", o[1]) if self.respond_to? o[0]
      end
    end
    # @see auto
    def auto=(v) Axlsx::validate_boolean v; @auto = v end    
    # @see rgb
    def rgb=(v) Axlsx::validate_string v; @rgb = v end    
    # @see tint
    def tint=(v) Axlsx::validate_float v; @tint = v end

    # This version does not support themes
    # def theme=(v) Axlsx::validate_unsigned_integer v; @theme = v end    
    
    # Indexed colors are for backward compatability which I am choosing not to support
    # def indexed=(v) Axlsx::validate_unsigned_integer v; @indexed = v end     


    # Serializes the color
    # @param [Nokogiri::XML::Builder] xml The document builder instance this objects xml will be added to.
    # @return [String]
    def to_xml(xml) xml.color(self.instance_values) end
  end
end
