# encoding: UTF-8
module Axlsx
  # The ValAxisData class manages the values for a chart value series.
  class ValAxisData < CatAxisData

    def to_xml_string(str = '')
      str << '<val>'
      str << '<numRef>'
      str << '<f>' << Axlsx::cell_range(@list) << '</f>'
      str << '<numCache>'
      str << '<formatCode>General</formatCode>'
      str << '<ptCount val="' << size.to_s << '"/>'
      each_with_index do |item, index|
        v = item.is_a?(Cell) ?  item.value.to_s : item
        str << '<pt idx="' << index.to_s << '"><v>' << v << '</v></pt>'
      end
      str << '</numCache>'
      str << '</numRef>'
      str << '</val>'
    end

    # Serializes the value axis data
    # @param [Nokogiri::XML::Builder] xml The document builder instance this objects xml will be added to.
    # @return [String]
    def to_xml(xml)
      xml.val {
        xml.numRef {
          xml.f Axlsx::cell_range(@list)
          xml.numCache {
            xml.formatCode 'General'
            xml.ptCount :val=>size
            each_with_index do |item, index|
              v = item.is_a?(Cell) ? item.value : item
              xml.pt(:idx=>index) { xml.v v }
            end
          }
        }
      }
    end

  end

end
