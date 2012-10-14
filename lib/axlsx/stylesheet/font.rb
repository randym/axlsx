# encoding: UTF-8
module Axlsx
  # The Font class details a font instance for use in styling cells.
  # @note The recommended way to manage fonts, and other styles is Styles#add_style
  # @see Styles#add_style
  class Font
    include Axlsx::OptionsParser

    # Creates a new Font
    # @option options [String] name
    # @option options [Integer] charset
    # @option options [Integer] family
    # @option options [Integer] family
    # @option options [Boolean] b
    # @option options [Boolean] i
    # @option options [Boolean] u
    # @option options [Boolean] strike
    # @option options [Boolean] outline
    # @option options [Boolean] shadow
    # @option options [Boolean] condense
    # @option options [Boolean] extend
    # @option options [Color] color
    # @option options [Integer] sz
    def initialize(options={})
      parse_options options
    end

    # The name of the font
    # @return [String]
    attr_reader :name

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
    attr_reader :charset

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
    attr_reader :family

    # Indicates if the font should be rendered in *bold*
    # @return [Boolean]
    attr_reader :b

    # Indicates if the font should be rendered italicized
    # @return [Boolean]
    attr_reader :i

    # Indicates if the font should be rendered underlined
    # @return [Boolean]
    attr_reader :u

    # Indicates if the font should be rendered with a strikthrough
    # @return [Boolean]
    attr_reader :strike

    # Indicates if the font should be rendered with an outline
    # @return [Boolean]
    attr_reader :outline

    # Indicates if the font should be rendered with a shadow
    # @return [Boolean]
    attr_reader :shadow

    # Indicates if the font should be condensed
    # @return [Boolean]
    attr_reader :condense

    # The font's extend property
    # @return [Boolean]
    attr_reader  :extend

    # The color of the font
    # @return [Color]
    attr_reader :color

    # The size of the font.
    # @return [Integer]
    attr_reader :sz

     # @see name
    def name=(v) Axlsx::validate_string v; @name = v end
    # @see charset
    def charset=(v) Axlsx::validate_unsigned_int v; @charset = v end
    # @see family
    def family=(v) Axlsx::validate_unsigned_int v; @family = v end
    # @see b
    def b=(v) Axlsx::validate_boolean v; @b = v end
    # @see i
    def i=(v) Axlsx::validate_boolean v; @i = v end
    # @see u
    def u=(v) Axlsx::validate_boolean v; @u = v end
    # @see strike
    def strike=(v) Axlsx::validate_boolean v; @strike = v end
    # @see outline
    def outline=(v) Axlsx::validate_boolean v; @outline = v end
    # @see shadow
    def shadow=(v) Axlsx::validate_boolean v; @shadow = v end
    # @see condense
    def condense=(v) Axlsx::validate_boolean v; @condense = v end
    # @see extend
    def extend=(v) Axlsx::validate_boolean v; @extend = v end
    # @see color
    def color=(v) DataTypeValidator.validate "Font.color", Color, v; @color=v end
    # @see sz
    def sz=(v) Axlsx::validate_unsigned_int v; @sz=v end

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      str << '<font>'
      instance_values.each do |k, v|
        v.is_a?(Color) ? v.to_xml_string(str) : (str << '<' << k.to_s << ' val="' << v.to_s << '"/>')
      end
      str << '</font>'
    end
  end
end
