module Axlsx
  # CellAlignment stores information about the cell alignment of a style Xf Object.
  # @note Using Styles#add_style is the recommended way to manage cell alignment. 
  # @see Styles#add_style
  class CellAlignment
    # The horizontal alignment of the cell.
    # @note 
    #  The horizontal cell alignement style must be one of
    #   :general
    #   :left
    #   :center
    #   :right
    #   :fill
    #   :justify
    #   :centerContinuous
    #   :distributed
    # @return [Symbol]
    attr_accessor :horizontal

    # The vertical alignment of the cell.
    # @note
    #  The vertical cell allingment style must be one of the following:
    #   :top
    #   :center
    #   :bottom
    #   :justify
    #   :distributed
    # @return [Symbol]
    attr_accessor :vertical

    # The textRotation of the cell.
    # @return [Integer]
    attr_accessor :textRotation

    # Indicate if the text of the cell should wrap
    # @return [Boolean]
    attr_accessor :wrapText
    
    # The amount of indent
    # @return [Integer]
    attr_accessor :indent

    # The amount of relativeIndent
    # @return [Integer]
    attr_accessor :relativeIndent

    # Indicate if the last line should be justified.
    # @return [Boolean]
    attr_accessor :justifyLastLine

    # Indicate if the text should be shrunk to the fit in the cell.
    # @return [Boolean]
    attr_accessor :shrinkToFit

    # The reading order of the text
    # 0 Context Dependent 
    # 1 Left-to-Right
    # 2 Right-to-Left
    # @return [Integer]
    attr_accessor :readingOrder
    
    # Create a new cell_alignment object
    # @option options [Symbol] horizontal
    # @option options [Symbol] vertical
    # @option options [Integer] textRotation
    # @option options [Boolean] wrapText
    # @option options [Integer] indent
    # @option options [Integer] relativeIndent
    # @option options [Boolean] justifyLastLine
    # @option options [Boolean] shrinkToFit
    # @option options [Integer] readingOrder
    def initialize(options={})
      options.each do |o|
        self.send("#{o[0]}=", o[1]) if self.respond_to? "#{o[0]}="
      end
    end        
    
    def horizontal=(v) Axlsx::validate_horizontal_alignment v; @horizontal = v end    
    def vertical=(v) Axlsx::validate_vertical_alignment v; @vertical = v end    
    def textRotation=(v) Axlsx::validate_unsigned_int v; @textRotation = v end    
    def wrapText=(v) Axlsx::validate_boolean v; @wrapText = v end    
    def indent=(v) Axlsx::validate_unsigned_int v; @indent = v end    
    def relativeIndent=(v) Axlsx::validate_int v; @relativeIndent = v end    
    def justifyLastLine=(v) Axlsx::validate_boolean v; @justifyLastLine = v end    
    def shrinkToFit=(v) Axlsx::validate_boolean v; @shrinkToFit = v end    
    def readingOrder=(v) Axlsx::validate_unsigned_int v; @readingOrder = v end  

    # Serializes the cell alignment
    # @param [Nokogiri::XML::Builder] xml The document builder instance this objects xml will be added to.
    # @return [String]
    def to_xml(xml)
      xml.alignment(self.instance_values)      
    end
    
  end
end
