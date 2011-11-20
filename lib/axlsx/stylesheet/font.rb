module Axlsx
  # The Font class details a font instance for use in styling cells.
  # @note The recommended way to manage fonts, and other styles is Styles#add_style
  # @see Styles#add_style
  class Font
    # The name of the font
    # @return [String]
    attr_accessor :name
    
    # The charset of the font
    # @return [Integer]
    # @note
    #  The following values are defined in the OOXML specification and are OS dependant values
    #   0   ANSI_CHARSET
    #   1   DEFAULT_CHARSET
    #   2   SYMBOL_CHARSET
    #   77  MAC_CHARSET
    #   128 SHIFTJIS_CHARSET
    #   129 HANGUL_CHARSET
    #   130 JOHAB_CHARSET
    #   134 GB2312_CHARSET
    #   136 CHINESEBIG5_CHARSET
    #   161 GREEK_CHARSET
    #   162 TURKISH_CHARSET
    #   163 VIETNAMESE_CHARSET
    #   177 HEBREW_CHARSET
    #   178 ARABIC_CHARSET
    #   186 BALTIC_CHARSET
    #   204 RUSSIAN_CHARSET
    #   222 THAI_CHARSET
    #   238 EASTEUROPE_CHARSET
    #   255 OEM_CHARSET
    attr_accessor :charset
    
    # The font's family
    # @note 
    #  The following are defined OOXML specification
    #   0 Not applicable.
    #   1 Roman
    #   2 Swiss
    #   3 Modern
    #   4 Script
    #   5 Decorative
    #   6..14 Reserved for future use
    # @return [Integer]
    attr_accessor :family

    # Indicates if the font should be rendered in *bold*
    # @return [Boolean]
    attr_accessor :b

    # Indicates if the font should be rendered italicized
    # @return [Boolean]
    attr_accessor :i

    # Indicates if the font should be rendered with a strikthrough
    # @return [Boolean]
    attr_accessor :strike

    # Indicates if the font should be rendered with an outline
    # @return [Boolean]
    attr_accessor :outline

    # Indicates if the font should be rendered with a shadow
    # @return [Boolean]
    attr_accessor :shadow

    # Indicates if the font should be condensed
    # @return [Boolean]
    attr_accessor :condense

    # The font's extend property
    # @return [Boolean]
    attr_accessor  :extend

    # The color of the font
    # @return [Color]
    attr_accessor :color

    # The size of the font.
    # @return [Integer]
    attr_accessor :sz

    # Creates a new Font
    # @option options [String] name
    # @option options [Integer] charset
    # @option options [Integer] family
    # @option options [Integer] family
    # @option options [Boolean] b
    # @option options [Boolean] i
    # @option options [Boolean] strike
    # @option options [Boolean] outline
    # @option options [Boolean] shadow
    # @option options [Boolean] condense
    # @option options [Boolean] extend
    # @option options [Color] color
    # @option options [Integer] sz
    def initialize(options={})
      options.each do |o|
        self.send("#{o[0]}=", o[1]) if self.respond_to? o[0]
      end
    end

    def name=(v) Axlsx::validate_string v; @name = v end    
    def charset=(v) Axlsx::validate_unsigned_int v; @charset = v end
    def family=(v) Axlsx::validate_unsigned_int v; @family = v end
    def b=(v) Axlsx::validate_boolean v; @b = v end    
    def i=(v) Axlsx::validate_boolean v; @i = v end
    def strike=(v) Axlsx::validate_boolean v; @strike = v end
    def outline=(v) Axlsx::validate_boolean v; @outline = v end
    def shadow=(v) Axlsx::validate_boolean v; @shadow = v end    
    def condense=(v) Axlsx::validate_boolean v; @condense = v end
    def extend=(v) Axlsx::validate_boolean v; @extend = v end
    def color=(v) DataTypeValidator.validate "Font.color", Color, v; @color=v end
    def sz=(v) Axlsx::validate_unsigned_int v; @sz=v end

    # Serializes the fill
    # @param [Nokogiri::XML::Builder] xml The document builder instance this objects xml will be added to.
    # @return [String]
    def to_xml(xml)
      xml.font {
        self.instance_values.each do |k, v|
          v.is_a?(Color) ? v.to_xml(xml) : xml.send(k, {:val => v})            
        end
      }
    end
  end
end
