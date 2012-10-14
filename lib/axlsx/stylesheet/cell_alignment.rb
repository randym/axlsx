# encoding: UTF-8
module Axlsx
  
 
  # CellAlignment stores information about the cell alignment of a style Xf Object.
  # @note Using Styles#add_style is the recommended way to manage cell alignment.
  # @see Styles#add_style
  class CellAlignment
    

    include Axlsx::SerializedAttributes
    include Axlsx::OptionsParser
    
    serializable_attributes :horizontal, :vertical, :text_rotation, :wrap_text, :indent, :relative_indent, :justify_last_line, :shrink_to_fit, :reading_order
    # Create a new cell_alignment object
    # @option options [Symbol] horizontal
    # @option options [Symbol] vertical
    # @option options [Integer] text_rotation
    # @option options [Boolean] wrap_text
    # @option options [Integer] indent
    # @option options [Integer] relative_indent
    # @option options [Boolean] justify_last_line
    # @option options [Boolean] shrink_to_fit
    # @option options [Integer] reading_order
    def initialize(options={})
      parse_options options
    end



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
    attr_reader :horizontal

    # The vertical alignment of the cell.
    # @note
    #  The vertical cell allingment style must be one of the following:
    #   :top
    #   :center
    #   :bottom
    #   :justify
    #   :distributed
    # @return [Symbol]
    attr_reader :vertical

    # The textRotation of the cell.
    # @return [Integer]
    attr_reader :text_rotation
    alias :textRotation :text_rotation

    # Indicate if the text of the cell should wrap
    # @return [Boolean]
    attr_reader :wrap_text
    alias :wrapText :wrap_text

    # The amount of indent
    # @return [Integer]
    attr_reader :indent

    # The amount of relativeIndent
    # @return [Integer]
    attr_reader :relative_indent
    alias :relativeIndent :relative_indent

    # Indicate if the last line should be justified.
    # @return [Boolean]
    attr_reader :justify_last_line
    alias :justifyLastLine :justify_last_line

    # Indicate if the text should be shrunk to the fit in the cell.
    # @return [Boolean]
    attr_reader :shrink_to_fit
    alias :shrinkToFit :shrink_to_fit

    # The reading order of the text
    # 0 Context Dependent
    # 1 Left-to-Right
    # 2 Right-to-Left
    # @return [Integer]
    attr_reader :reading_order
    alias :readingOrder :reading_order

    # @see horizontal
    def horizontal=(v) Axlsx::validate_horizontal_alignment v; @horizontal = v end
    # @see vertical
    def vertical=(v) Axlsx::validate_vertical_alignment v; @vertical = v end
    # @see textRotation
    def text_rotation=(v) Axlsx::validate_unsigned_int v; @text_rotation = v end
    alias :textRotation= :text_rotation=

    # @see wrapText
    def wrap_text=(v) Axlsx::validate_boolean v; @wrap_text = v end
    alias :wrapText= :wrap_text=

    # @see indent
    def indent=(v) Axlsx::validate_unsigned_int v; @indent = v end

    # @see relativeIndent
    def relative_indent=(v) Axlsx::validate_int v; @relative_indent = v end
    alias :relativeIndent= :relative_indent=

    # @see justifyLastLine
    def justify_last_line=(v) Axlsx::validate_boolean v; @justify_last_line = v end
    alias :justifyLastLine= :justify_last_line=

    # @see shrinkToFit
    def shrink_to_fit=(v) Axlsx::validate_boolean v; @shrink_to_fit = v end
    alias :shrinkToFit= :shrink_to_fit=

    # @see readingOrder
    def reading_order=(v) Axlsx::validate_unsigned_int v; @reading_order = v end
    alias :readingOrder= :reading_order=

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      str << '<alignment '
      serialized_attributes str
      str << '/>'
    end

  end
end
