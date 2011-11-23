module Axlsx
  # The CatAxisData class serializes the category axis data for a chart
  class CatAxisData < SimpleTypedList

    # Create a new CatAxisData object
    # @param [Array, SimpleTypedList] data the data for this category axis. This can be a simple array or a simple typed list of cells.
    def initialize(data=[])
      super Object
      @list.concat data if data.is_a?(Array)
    end

    # Serializes the category axis data
    # @param [Nokogiri::XML::Builder] xml The document builder instance this objects xml will be added to.
    # @return [String]
    def to_xml(xml)
      xml.send('c:cat') {
        xml.send('c:strRef') {
          xml.send('c:f', Axlsx::cell_range(@list))
          xml.send('c:strCache') {
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
