# encoding: UTF-8
module Axlsx
  # The CatAxisData class serializes the category axis data for a chart
  class CatAxisData < SimpleTypedList

    # Create a new CatAxisData object
    # @param [Array, SimpleTypedList] data the data for this category axis. This can be a simple array or a simple typed list of cells.
    def initialize(data=[])
      super Object
      @list.concat data if data.is_a?(Array)
      data.each { |i| @list << i } if data.is_a?(SimpleTypedList)
    end


    def to_xml_string(str = '')
      str << '<c:cat>'
      str << '<c:strRef>'
      str << '<c:f>' << Axlsx::cell_range(@list) << '</c:f>'
      str << '<c:strCache>'
      str << '<c:ptCount val="' << size.to_s << '"/>'
      each_with_index do |item, index|
        v = item.is_a?(Cell) ?  item.value.to_s : item
        str << '<c:pt idx="' << index.to_s << '"><c:v>' << v << '</c:v></c:pt>'
      end
      str << '</c:strCache>'
      str << '</c:strRef>'
      str << '</c:cat>'
    end

  end

end
