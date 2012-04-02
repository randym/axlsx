# encoding: UTF-8
module Axlsx
  # A series title is a Title with a slightly different serialization than chart titles.
  class SeriesTitle < Title

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      str << '<c:tx>'
      str << '<c:strRef>'
      str << '<c:f>' << Axlsx::cell_range([@cell]) << '</c:f>'
      str << '<c:strCache>'
      str << '<c:ptCount val="1"/>'
      str << '<c:pt idx="0">'
      str << '<c:v>' << @text << '</c:v>'
      str << '</c:pt>'
      str << '</c:strCache>'
      str << '</c:strRef>'
      str << '</c:tx>'
    end
  end
end
