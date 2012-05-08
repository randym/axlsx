module Axlsx

  class VmlDrawing

    def initialize(comments)
      raise ArgumentError, "you must provide a comments object" unless comments.is_a?(Comments)
      @comments = comments
    end

    def pn
      "#{VML_DRAWING_PN}" % (@comments.worksheet.index + 1)
    end

    def to_xml_string(str = '')
      str = <<BAD_PROGRAMMER
<xml xmlns:v="urn:schemas-microsoft-com:vml"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel">
 <o:shapelayout v:ext="edit">
  <o:idmap v:ext="edit" data="#{@comments.worksheet.index+1}"/>
 </o:shapelayout>
 <v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202"
  path="m0,0l0,21600,21600,21600,21600,0xe">
  <v:stroke joinstyle="miter"/>
  <v:path gradientshapeok="t" o:connecttype="rect"/>
 </v:shapetype>
BAD_PROGRAMMER
      @comments.comment_list.each { |comment| comment.vml_shape.to_xml_string str }
      str << "</xml>"

    end

  end
end
