module Axlsx

  # A VmlShape is used to position and render a comment.
  class VmlShape

    include Axlsx::OptionsParser
    include Axlsx::Accessors

    # Creates a new VmlShape
    # @option options [Integer] row
    # @option options [Integer] column
    # @option options [Integer] left_column
    # @option options [Integer] left_offset
    # @option options [Integer] top_row
    # @option options [Integer] top_offset
    # @option options [Integer] right_column
    # @option options [Integer] right_offset
    # @option options [Integer] bottom_row
    # @option options [Integer] bottom_offset
    def initialize(options={})
      @row = @column = @left_column = @top_row = @right_column = @bottom_row = 0
      @left_offset = 15
      @top_offset = 2
      @right_offset = 50
      @bottom_offset = 5
      @visible = true
      @id = (0...8).map{65.+(rand(25)).chr}.join
      parse_options options
      yield self if block_given?
    end

    unsigned_int_attr_accessor :row, :column, :left_column, :left_offset, :top_row, :top_offset,
                               :right_column, :right_offset, :bottom_row, :bottom_offset

    boolean_attr_accessor :visible

    # serialize the shape to a string
    # @param [String] str
    # @return [String]
    def to_xml_string(str ='')
str << <<SHAME_ON_YOU

<v:shape id="#{@id}" type="#_x0000_t202" fillcolor="#ffffa1 [80]" o:insetmode="auto"
  style="visibility:#{@visible ? 'visible' : 'hidden'}">
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
   #{@visible ? '<x:Visible/>' : ''}
  </x:ClientData>
 </v:shape>
SHAME_ON_YOU

    end
  end
end
