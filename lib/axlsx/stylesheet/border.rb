# encoding: UTF-8
module Axlsx
  # This class details a border used in Office Open XML spreadsheet styles.
  class Border

    include Axlsx::SerializedAttributes
    include Axlsx::OptionsParser

    # Creates a new Border object
    # @option options [Boolean] diagonal_up
    # @option options [Boolean] diagonal_down
    # @option options [Boolean] outline
    # @example - Making a border
    #   p = Axlsx::Package.new
    #   red_border = p.workbook.styles.add_style :border => { :style => :thin, :color => "FFFF0000" }
    #   ws = p.workbook.add_worksheet
    #   ws.add_row [1,2,3], :style => red_border
    #   p.serialize('red_border.xlsx')
    #
    # @note The recommended way to manage borders is with Style#add_style
    # @see Style#add_style
    def initialize(options={})
      @prs = SimpleTypedList.new BorderPr
      parse_options options
    end

    serializable_attributes :diagonal_up, :diagonal_down, :outline

    # @return [Boolean] The diagonal up property for the border that indicates if the border should include a diagonal line from the bottom left to the top right of the cell.
    attr_reader :diagonal_up
    alias :diagonalUp :diagonal_up

    # @return [Boolean] The diagonal down property for the border that indicates if the border should include a diagonal line from the top left to the top right of the cell.
    attr_reader :diagonal_down
    alias :diagonalDown :diagonal_down

    # @return [Boolean] The outline property for the border indicating that top, left, right and bottom borders should only be applied to the outside border of a range of cells.
    attr_reader :outline

    # @return [SimpleTypedList] A list of BorderPr objects for this border.
    attr_reader :prs

    # @see diagonalUp
    def diagonal_up=(v) Axlsx::validate_boolean v; @diagonal_up = v end
    alias :diagonalUp= :diagonal_up=

    # @see diagonalDown
    def diagonal_down=(v) Axlsx::validate_boolean v; @diagonal_down = v end
    alias :diagonalDown= :diagonal_down=

    # @see outline
    def outline=(v) Axlsx::validate_boolean v; @outline = v end

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      str << '<border '
      serialized_attributes str
      str << '>'
      # enforces order
      [:start, :end, :left, :right, :top, :bottom, :diagonal, :vertical, :horizontal].each do |k|
        @prs.select { |pr| pr.name == k }.each do |part|
          part.to_xml_string(str)
        end
      end
      str << '</border>'
    end

  end
end
