module Axlsx

  # A VmlShape is used to position and render a comment.
  class VmlShape

    # The row anchor position for this shape determined by the comment's ref value
    # @return [Integer]
    attr_reader :row

    # The column anchor position for this shape determined by the comment's ref value
    # @return [Integer]
    attr_reader :column

    # The left column for this shape
    # @return [Integer]
    attr_reader :left_column

    # The left offset for this shape
    # @return [Integer]
    attr_reader :left_offset

    # The top row for this shape
    # @return [Integer]
    attr_reader :top_row

    # The top offset for this shape
    # @return [Integer]
    attr_reader :top_offset

    # The right column for this shape
    # @return [Integer]
    attr_reader :right_column

    # The right offset for this shape
    # @return [Integer]
    attr_reader :right_offset

    # The botttom row for this shape
    # @return [Integer]
    attr_reader :bottom_row

    # The bottom offset for this shape
    # @return [Integer]
    attr_reader :bottom_offset

    # Creates a new VmlShape
    # @option options [Integer|String] left_column
    # @option options [Integer|String] left_offset
    # @option options [Integer|String] top_row
    # @option options [Integer|String] top_offset
    # @option options [Integer|String] right_column
    # @option options [Integer|String] right_offset
    # @option options [Integer|String] bottom_row
    # @option options [Integer|String] bottom_offset
    def initialize(options={})
      @row = @column = @left_column = @top_row = @right_column = @bottom_row = 0
      @left_offset = 15
      @top_offset = 2
      @right_offset = 50
      @bottom_offset = 5
      @id = (0...8).map{65.+(rand(25)).chr}.join
      options.each do |o|
        self.send("#{o[0]}=", o[1]) if self.respond_to? "#{o[0]}="
      end
      yield self if block_given?
    end

    # @see column
    def column=(v); Axlsx::validate_integerish(v); @column = v.to_i end
 
    # @see row
    def row=(v); Axlsx::validate_integerish(v); @row = v.to_i end
    # @see left_column
    def left_column=(v); Axlsx::validate_integerish(v); @left_column = v.to_i end

    # @see left_offset
    def left_offset=(v); Axlsx::validate_integerish(v); @left_offset = v.to_i end

    # @see top_row
    def top_row=(v); Axlsx::validate_integerish(v); @top_row = v.to_i end

    # @see top_offset
    def top_offset=(v); Axlsx::validate_integerish(v); @top_offset = v.to_i end

    # @see right_column
    def right_column=(v); Axlsx::validate_integerish(v); @right_column = v.to_i end

    # @see right_offset
    def right_offset=(v); Axlsx::validate_integerish(v); @right_offset = v.to_i end

    # @see bottom_row
    def bottom_row=(v); Axlsx::validate_integerish(v); @bottom_row = v.to_i end

    # @see bottom_offset
    def bottom_offset=(v); Axlsx::validate_integerish(v); @bottom_offset = v.to_i end

    # serialize the shape to a string
    # @param [String] str
    # @return [String]
    def to_xml_string(str ='')
str << <<SHAME_ON_YOU

<v:shape id="#{@id}" type="#_x0000_t202" fillcolor="#ffffa1 [80]" o:insetmode="auto">
  <v:fill color2="#ffffa1 [80]"/>
  <v:shadow on="t" obscured="t"/>
  <v:path o:connecttype="none"/>
  <v:textbox style='mso-fit-text-with-word-wrap:t'>
   <div style='text-align:left'></div>
  </v:textbox>

  <x:ClientData ObjectType="Note">
   <x:MoveWithCells/>
   <x:SizeWithCells/>
   <x:Anchor>#{left_column}, #{left_offset}, #{top_row}, #{top_offset}, #{right_column}, #{right_offset}, #{bottom_row}, #{bottom_offset}</x:Anchor>
   <x:AutoFill>False</x:AutoFill>
   <x:Row>#{row}</x:Row>
   <x:Column>#{column}</x:Column>
   <x:Visible/>
  </x:ClientData>
 </v:shape>
SHAME_ON_YOU

    end
  end
end
