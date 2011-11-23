module Axlsx
  # A graphic frame defines a container for a chart object
  # @note The recommended way to manage charts is Worksheet#add_chart
  # @see Worksheet#add_chart
  class GraphicFrame

    # A reference to the chart object associated with this frame
    # @return [Chart]
    attr_reader :chart

    # A anchor that holds this frame
    # @return [TwoCellAnchor]
    attr_reader :anchor

    # The relationship id for this graphic
    # @return [String]
    attr_reader :rId

    # Creates a new GraphicFrame object
    # @param [TwoCellAnchor] anchor
    # @param [Class] chart_type
    def initialize(anchor, chart_type, options)
      DataTypeValidator.validate "Drawing.chart_type", Chart, chart_type 
      @anchor = anchor
      @chart = chart_type.new(self, options)
    end

    def rId 
      "rId#{@anchor.index+1}"
    end

    # Serializes the graphic frame
    # @param [Nokogiri::XML::Builder] xml The document builder instance this objects xml will be added to.
    # @return [String]
    def to_xml(xml)
      xml.send('xdr:graphicFrame') {        
        xml.send('xdr:nvGraphicFramePr') {
          xml.send('xdr:cNvPr', :id=>2, :name=>chart.title)
          xml.send('xdr:cNvGraphicFramePr')                
        }
        xml.send('xdr:xfrm') {
          xml.send('a:off', :x=>0, :y=>0)
          xml.send('a:ext', :cx=>0, :cy=>0)
        }
        xml.send('a:graphic') {
          xml.send('a:graphicData', :uri=>XML_NS_C) {
            xml.send('c:chart', :'xmlns:c'=>XML_NS_C, :'xmlns:r'=>XML_NS_R, :'r:id'=>rId)
          }
        }
      }
      
    end
  end
end
