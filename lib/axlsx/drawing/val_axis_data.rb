module Axlsx
  # The ValAxisData class manages the values for a chart value series.
  class ValAxisData < CatAxisData

    # Serializes the value axis data
    # @param [Nokogiri::XML::Builder] xml The document builder instance this objects xml will be added to.
    # @return [String]
    def to_xml(xml)
      xml.send('c:val') {
        xml.send('c:numRef') {
          xml.send('c:f', Axlsx::cell_range(@list))
          xml.send('c:numCache') {
            xml.send('c:formatCode', 'General')
            xml.send('c:ptCount', :val=>size)
            each_with_index do |item, index|
              v = item.is_a?(Cell) ? item.value : item
              xml.send('c:pt', :idx=>index) {
                xml.send('c:v', v) 
              }
            end
          }                        
        }
      }
    end

  end
  
end
