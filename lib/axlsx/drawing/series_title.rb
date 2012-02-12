# encoding: UTF-8
module Axlsx
  # A series title is a Title with a slightly different serialization than chart titles.
  class SeriesTitle < Title

    # Serializes the series title
    # @param [Nokogiri::XML::Builder] xml The document builder instance this objects xml will be added to.
    # @return [String]
    def to_xml(xml)
      xml[:c].tx {
        xml[:c].strRef {
          xml[:c].f Axlsx::cell_range([@cell])
          xml[:c].strCache {
            xml[:c].ptCount :val=>1
            xml[:c].pt(:idx=>0) {
              xml[:c].v @text
            }
          }
        }
      }
    end
  end
end
