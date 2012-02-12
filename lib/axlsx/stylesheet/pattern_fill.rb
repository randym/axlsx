# encoding: UTF-8
module Axlsx
  # A PatternFill is the pattern and solid fill styling for a cell.
  # @note The recommended way to manage styles is with Styles#add_style
  # @see Style#add_style
  class PatternFill

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

    # Creates a new PatternFill Object
    # @option options [Symbol] patternType
    # @option options [Color] fgColor
    # @option options [Color] bgColor
    def initialize(options={})
      @patternType = :none
      options.each do |o|
        self.send("#{o[0]}=", o[1]) if self.respond_to? "#{o[0]}="
      end
    end
    # @see fgColor
    def fgColor=(v) DataTypeValidator.validate "PatternFill.fgColor", Color, v; @fgColor=v end
    # @see bgColor
    def bgColor=(v) DataTypeValidator.validate "PatternFill.bgColor", Color, v; @bgColor=v end
    # @see patternType
    def patternType=(v) Axlsx::validate_pattern_type v; @patternType = v end    

    # Serializes the pattern fill
    # @param [Nokogiri::XML::Builder] xml The document builder instance this objects xml will be added to.
    # @return [String]
    def to_xml(xml) 
      xml.patternFill(:patternType => self.patternType) { 
        self.instance_values.reject { |k,v| k.to_sym == :patternType }.each { |k,v| xml.send(k, v.instance_values) }    
      }
    end
  end
end
