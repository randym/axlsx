# encoding: UTF-8
module Axlsx
  # The GradientStop object represents a color point in a gradient.
  # @see Open Office XML Part 1 §18.8.24
  class GradientStop
    # The color for this gradient stop
    # @return [Color]
    # @see Color
    attr_reader :color

    # The position of the color
    # @return [Float]
    attr_reader :position

    # Creates a new GradientStop object
    # @param [Color] color
    # @param [Float] position
    def initialize(color, position)
      self.color = color
      self.position = position
    end

    # @see color
    def color=(v) DataTypeValidator.validate "GradientStop.color", Color, v; @color=v end
    # @see position
    def position=(v) DataTypeValidator.validate "GradientStop.position", Float, v, lambda { |arg| arg >= 0 && arg <= 1}; @position = v end

    def to_xml_string(str = '')
      str << '<stop position="' << position.to_s << '">'
      self.color.to_xml_string(str)
      str << '</stop>'
    end
    # Serializes the gradientStop
    # @param [Nokogiri::XML::Builder] xml The document builder instance this objects xml will be added to.
    # @return [String]
    def to_xml(xml) xml.stop(:position => self.position) {self.color.to_xml(xml)} end
  end
end
