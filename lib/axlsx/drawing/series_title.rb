module Axlsx
  # A series title is a Title with a slightly different serialization than chart titles.
  class SeriesTitle < Title

    # Serializes the series title
    # @param [Nokogiri::XML::Builder] xml The document builder instance this objects xml will be added to.
    # @return [String]
    def to_xml(xml)
      xml.send('c:tx') {
        xml.send('c:strRef') {
          xml.send('c:f', range)
          xml.send('c:strCache') {
            xml.send('c:ptCount', :val=>1)
            xml.send('c:pt', :idx=>0) {
              xml.send('c:v', @text)
            }
          }
        }
      }
    end
  end
end
