# encoding: UTF-8
module Axlsx
  # A PatternFill is the pattern and solid fill styling for a cell.
  # @note The recommended way to manage styles is with Styles#add_style
  # @see Style#add_style
  class PatternFill

    include Axlsx::OptionsParser
    # Creates a new PatternFill Object
    # @option options [Symbol] patternType
    # @option options [Color] fgColor
    # @option options [Color] bgColor
    def initialize(options={})
      @patternType = :none
      parse_options options
    end

    # The color to use for the the background in solid fills.
    # @return [Color]
    attr_reader :fgColor

    # The color to use for the background of the fill when the type is not solid.
    # @return [Color]
    attr_reader :bgColor

    # The pattern type to use
    # @note
    #  patternType must be one of
    #   :none
    #   :solid
    #   :mediumGray
    #   :darkGray
    #   :lightGray
    #   :darkHorizontal
    #   :darkVertical
    #   :darkDown
    #   :darkUp
    #   :darkGrid
    #   :darkTrellis
    #   :lightHorizontal
    #   :lightVertical
    #   :lightDown
    #   :lightUp
    #   :lightGrid
    #   :lightTrellis
    #   :gray125
    #   :gray0625
    # @see Office Open XML Part 1 18.18.55
    attr_reader :patternType

    # @see fgColor
    def fgColor=(v) DataTypeValidator.validate "PatternFill.fgColor", Color, v; @fgColor=v end
    # @see bgColor
    def bgColor=(v) DataTypeValidator.validate "PatternFill.bgColor", Color, v; @bgColor=v end
    # @see patternType
    def patternType=(v) Axlsx::validate_pattern_type v; @patternType = v end

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      str << ('<patternFill patternType="' << patternType.to_s << '">')
      if fgColor.is_a?(Color)
        fgColor.to_xml_string str, "fgColor"
      end

      if bgColor.is_a?(Color)
        bgColor.to_xml_string str, "bgColor"
      end
      str << '</patternFill>'
    end
  end
end
