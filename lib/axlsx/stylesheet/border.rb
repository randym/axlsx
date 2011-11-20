module Axlsx
  # This class details a border used in Office Open XML spreadsheet styles.
  class Border

    # @return [Boolean] The diagonal up property for the border that indicates if the border should include a diagonal line from the bottom left to the top right of the cell. 
    attr_accessor :diagonalUp
    
    # @return [Boolean] The diagonal down property for the border that indicates if the border should include a diagonal line from the top left to the top right of the cell.
    attr_accessor :diagonalDown

    # @return [Boolean] The outline property for the border indicating that top, left, right and bottom borders should only be applied to the outside border of a range of cells.
    attr_accessor :outline

    # @return [SimpleTypedList] A list of BorderPr objects for this border. 
    attr_reader :prs

    # Creates a new Border object
    # @option options [Boolean] diagonalUp
    # @option options [Boolean] diagonalDown
    # @option options [Boolean] outline
    # @example Making a border
    #   p = Package.new
    #   red_border = Border.new
    #   [:left, :right, :top, :bottom].each do |item| 
    #     red_border.prs << BorderPr.new(:name=>item, :style=>:thin, :color=>Color.new(:rgb=>"FFFF0000"))     #   
    #   end
    #   # this sets red_border to be the index for the created border.
    #   red_border = p.workbook.styles.@borders << red_border
    #   #used in row creation as follows. This will add a red border to each of the cells in the row.
    #   p.workbook.add_worksheet.rows << :values=>[1,2,3] :style=>red_border
    def initialize(options={})
      @prs = SimpleTypedList.new BorderPr
      options.each do |o|
        self.send("#{o[0]}=", o[1]) if self.respond_to? "#{o[0]}="
      end
    end        

    def diagonalUp=(v) Axlsx::validate_boolean v; @diagonalUp = v end
    def diagonalDown=(v) Axlsx::validate_boolean v; @diagonalDown = v end
    def outline=(v) Axlsx::validate_boolean v; @outline = v end

    # Serializes the border element
    # @param [Nokogiri::XML::Builder] xml The document builder instance this objects xml will be added to.
    def to_xml(xml)
      xml.border(self.instance_values.select{ |k,v| [:diagonalUp, :diagonalDown, :outline].include? k }) {
        [:start, :end, :left, :right, :top, :bottom, :diagonal, :vertical, :horizontal].each do |k|
          @prs.select { |pr| pr.name == k }.each { |pr| pr.to_xml(xml) }
        end
      }
    end
  end
end
