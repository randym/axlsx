# encoding: UTF-8
module Axlsx
  # A border part.
  class BorderPr
    include Axlsx::OptionsParser
    # @return [Color] The color of this border part.
    attr_reader :color

    # @return [Symbol] The syle of this border part.
    # @note
    #  The following are allowed
    #   :none
    #   :thin
    #   :medium
    #   :dashed
    #   :dotted
    #   :thick
    #   :double
    #   :hair
    #   :mediumDashed
    #   :dashDot
    #   :mediumDashDot
    #   :dashDotDot
    #   :mediumDashDotDot
    #   :slantDashDot
    attr_reader :style

    # @return [Symbol] The name of this border part
    # @note
    #  The following are allowed
    #   :start
    #   :end
    #   :left
    #   :right
    #   :top
    #   :bottom
    #   :diagonal
    #   :vertical
    #   :horizontal
    attr_reader :name

    # Creates a new Border Part Object
    # @option options [Color] color
    # @option options [Symbol] name
    # @option options [Symbol] style
    # @see Axlsx::Border
    def initialize(options={})
      parse_options(options)
      #options.each do |o|
      #  self.send("#{o[0]}=", o[1]) if self.respond_to? "#{o[0]}="
      #end
    end

    # @see name
    def name=(v) RestrictionValidator.validate "BorderPr.name", [:start, :end, :left, :right, :top, :bottom, :diagonal, :vertical, :horizontal], v; @name = v end
    # @see color
    def color=(v) DataTypeValidator.validate(:color, Color, v); @color = v end
    # @see style
    def style=(v) RestrictionValidator.validate "BorderPr.style", [:none, :thin, :medium, :dashed, :dotted, :thick, :double, :hair, :mediumDashed, :dashDot, :mediumDashDot, :dashDotDot, :mediumDashDotDot, :slantDashDot], v; @style = v end

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      str << '<' << @name.to_s << ' style="' << @style.to_s << '">'
      @color.to_xml_string(str) if @color.is_a?(Color)
      str << '</' << @name.to_s << '>'
    end

  end
end
