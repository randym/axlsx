# encoding: UTF-8
module Axlsx
  # The ValAxisData class manages the values for a chart value series.
  class NamedAxisData < CatAxisData

    def initialize(name, data=[])
      super(data)
      @name = name
    end

    # Serializes the value axis data
    # @param [Nokogiri::XML::Builder] xml The document builder instance this objects xml will be added to.
    # @return [String]
    def to_xml(xml)
      xml.send(@name) {
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
