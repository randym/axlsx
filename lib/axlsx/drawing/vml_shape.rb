module Axlsx

  class VmlShape

    attr_accessor :row

    attr_accessor :column

    attr_accessor :left_column
    attr_accessor :left_offset
    attr_accessor :top_row
    attr_accessor :top_offset
    attr_accessor :right_column
    attr_accessor :right_offset
    attr_accessor :bottom_row
    attr_accessor :bottom_offset
    attr_reader   :id

    def initialize(comment, options={})
      @id = "_x0000_s#{comment.comments.worksheet.index+1}08#{comment.index+1}"
      @row = @column = @left_column = @top_row = @right_column = @bottom_row = 0
      @left_offset = 15
      @top_offset = 2
      @right_offset = 50
      @bottom_offset = 5
      options.each do |o|
        self.send("#{o[0]}=", o[1]) if self.respond_to? "#{o[0]}="
      end
      yield self if block_given?
    end

    def to_xml_string(str ='')
str << <<SHAME_ON_YOU

<v:shape id="#{id}" type="#_x0000_t202"
style='position:absolute;margin-left:104pt;margin-top:2pt;width:800px;height:27pt;z-index:1;mso-wrap-style:tight'
 fillcolor="#ffffa1 [80]" o:insetmode="auto">

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
