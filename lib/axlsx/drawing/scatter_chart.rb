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

    # Serializes the bar chart
    # @return [String]
    def to_xml
      super() do |xml|
        xml.scatterChart {
          xml.scatterStyle :val=>scatterStyle

          # This is all repeated from line_3D_chart.rb!
          xml.varyColors :val=>1
          @series.each { |ser| ser.to_xml(xml) }
          xml.dLbls {
            xml.showLegendKey :val=>0
            xml.showVal :val=>0
            xml.showCatName :val=>0
            xml.showSerName :val=>0
            xml.showPercent :val=>0
            xml.showBubbleSize :val=>0
          }
          xml.axId :val=>@xValAxId
          xml.axId :val=>@yValAxId
        }
        @xValAxis.to_xml(xml)
        @yValAxis.to_xml(xml)
      end
    end
  end
end
