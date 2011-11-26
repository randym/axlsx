module Axlsx
  # A border part. 
  class BorderPr
    
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
      options.each do |o|
        self.send("#{o[0]}=", o[1]) if self.respond_to? "#{o[0]}="
      end
    end

    # @see name
    def name=(v) RestrictionValidator.validate "BorderPr.name", [:start, :end, :left, :right, :top, :bottom, :diagonal, :vertical, :horizontal], v; @name = v end
    # @see color
    def color=(v) DataTypeValidator.validate(:color, Color, v); @color = v end    
    # @see style
    def style=(v) RestrictionValidator.validate "BorderPr.style", [:none, :thin, :medium, :dashed, :dotted, :thick, :double, :hair, :mediumDashed, :dashDot, :mediumDashDot, :dashDotDot, :mediumDashDotDot, :slantDashDot], v; @style = v end

    # Serializes the border part
    # @param [Nokogiri::XML::Builder] xml The document builder instance this objects xml will be added to.
    # @return [String]
    def to_xml(xml)
      xml.send(@name, :style => @style) {
        @color.to_xml(xml) if @color.is_a? Color
      } 
    end
  end
end
