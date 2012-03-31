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
      str << '<cat>'
      str << '<strRef>'
      str << '<f>' << Axlsx::cell_range(@list) << '</f>'
      str << '<strCache>'
      str << '<ptCount val="' << size.to_s << '"/>'
      each_with_index do |item, index|
        v = item.is_a?(Cell) ?  item.value.to_s : item
        str << '<pt idx="' << index.to_s << '"><v>' << v << '</v></pt>'
      end
      str << '</strCache>'
      str << '</strRef>'
      str << '</cat>'
    end

    # Serializes the category axis data
    # @param [Nokogiri::XML::Builder] xml The document builder instance this objects xml will be added to.
    # @return [String]
    def to_xml(xml)
      xml.cat {
        xml.strRef {
          xml.f Axlsx::cell_range(@list)
          xml.strCache {
            xml.ptCount :val=>size
            each_with_index do |item, index|
              v = item.is_a?(Cell) ? item.value : item
              xml.pt(:idx=>index) {
                xml.v v
              }
            end
          }
        }
      }
    end

  end

end
