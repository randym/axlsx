# encoding: UTF-8
module Axlsx
  class ScatterChart < Chart
    attr_reader :scatterStyle

    # the x value axis
    # @return [ValAxis]
    attr_reader :xValAxis

    # the y value axis
    # @return [ValAxis]
    attr_reader :yValAxis

    def initialize(frame, options={})
      @scatterStyle = :lineMarker
      @xValAxId = rand(8 ** 8)
      @yValAxId = rand(8 ** 8)
      @xValAxis = ValAxis.new(@xValAxId, @yValAxId)
      @yValAxis = ValAxis.new(@yValAxId, @xValAxId)
      super(frame, options)
      @series_type = ScatterSeries
    end

    def to_xml_string(str = '')
      super do |str|
        str << '<c:scatterChart>'
        str << '<c:scatterStyle val="' << scatterStyle.to_s << '"/>'
        str << '<c:varyColors val="1"/>'
        @series.each { |ser| ser.to_xml_string(str) }
        str << '<c:dLbls>'
        str << '<c:showLegendKey val="0"/>'
        str << '<c:showVal val="0"/>'
        str << '<c:showCatName val="0"/>'
        str << '<c:showSerName val="0"/>'
        str << '<c:showPercent val="0"/>'
        str << '<c:showBubbleSize val="0"/>'
        str << '</c:dLbls>'
        str << '<c:axId val="' << @xValAxId.to_s << '"/>'
        str << '<c:axId val="' << @yValAxId.to_s << '"/>'
        str << '</c:scatterChart>'
        @xValAxis.to_xml_string str
        @yValAxis.to_xml_string str
      end
      str
    end
  end
end
