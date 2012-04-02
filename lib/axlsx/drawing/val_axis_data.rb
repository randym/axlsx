# encoding: UTF-8
module Axlsx
  # The ValAxisData class manages the values for a chart value series.
  class ValAxisData < CatAxisData

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      str << '<c:val>'
      str << '<c:numRef>'
      str << '<c:f>' << Axlsx::cell_range(@list) << '</c:f>'
      str << '<c:numCache>'
      str << '<c:formatCode>General</c:formatCode>'
      str << '<c:ptCount val="' << size.to_s << '"/>'
      each_with_index do |item, index|
        v = item.is_a?(Cell) ?  item.value.to_s : item
        str << '<c:pt idx="' << index.to_s << '"><c:v>' << v << '</c:v></c:pt>'
      end
      str << '</c:numCache>'
      str << '</c:numRef>'
      str << '</c:val>'
    end

  end

end
