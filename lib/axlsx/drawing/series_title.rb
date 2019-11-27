# encoding: UTF-8
# frozen_string_literal: true
module Axlsx
  # A series title is a Title with a slightly different serialization than chart titles.
  class SeriesTitle < Title

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = String.new)
      str << '<c:tx>'\
             '<c:strRef>'\
             "<c:f>#{Axlsx::cell_range([@cell])}</c:f>"\
             '<c:strCache>'\
             '<c:ptCount val="1"/>'\
             '<c:pt idx="0">'\
             "<c:v>#{@text}</c:v>"\
             '</c:pt>'\
             '</c:strCache>'\
             '</c:strRef>'\
             '</c:tx>'
    end
  end
end
