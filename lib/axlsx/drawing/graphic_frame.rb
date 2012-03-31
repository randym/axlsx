# encoding: UTF-8
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

    # Creates a new GraphicFrame object
    # @param [TwoCellAnchor] anchor
    # @param [Class] chart_type
    def initialize(anchor, chart_type, options)
      DataTypeValidator.validate "Drawing.chart_type", Chart, chart_type
      @anchor = anchor
      @chart = chart_type.new(self, options)
    end

    # The relationship id for this graphic
    # @return [String]
    def rId
      "rId#{@anchor.index+1}"
    end

    def to_xml_string(str = '')
      str << '<graphicFrame>'
      str << '<nvGraphicFramePr>'
      str << '<cNvPr id="2" name="' << chart.title.text << '"/>'
      str << '<cNvGraphicFramePr/>'
      str << '</nvGraphicFramePr>'
      str << '<xfrm>'
      str << '<a:off x="0" y="0"/>'
      str << '<a:ext cx="0" cy="0"/>'
      str << '</xfrm>'
      str << '<a:graphic>'
      str << '<graphicData uri="' << XML_NS_C << '">'
      str << '<c:chart xmlns:c="' << XML_NS_C << '" xmlns:r="' << XML_NS_R << '" r:id="' << rId.to_s << '"/>'
      str << '</graphicData>'
      str << '</a:graphic>'
      str << '</graphicFrame>'
    end

    # Serializes the graphic frame
    # @param [Nokogiri::XML::Builder] xml The document builder instance this objects xml will be added to.
    # @return [String]
    def to_xml(xml)
      xml.graphicFrame {
        xml.nvGraphicFramePr {
          xml.cNvPr :id=>2, :name=>chart.title.text
          xml.cNvGraphicFramePr
        }
        xml.xfrm {
          xml[:a].off(:x=>0, :y=>0)
          xml[:a].ext :cx=>0, :cy=>0
        }
        xml[:a].graphic {
          xml.graphicData(:uri=>XML_NS_C) {
            xml[:c].chart :'xmlns:c'=>XML_NS_C, :'xmlns:r'=>XML_NS_R, :'r:id'=>rId
          }
        }
      }

    end
  end
end
