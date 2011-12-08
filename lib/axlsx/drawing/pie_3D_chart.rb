module Axlsx


  # The Pie3DChart is a three dimentional piechart (who would have guessed?) that you can add to your worksheet.
  # @see Worksheet#add_chart
  # @see Chart#add_series
  # @see README for an example
  class Pie3DChart < Chart

    # Creates a new pie chart object
    # @param [GraphicFrame] frame The workbook that owns this chart.
    # @option options [Cell, String] title
    # @option options [Boolean] show_legend    
    # @option options [Symbol] grouping
    # @option options [String] gapDepth
    # @option options [Integer] rotX
    # @option options [String] hPercent
    # @option options [Integer] rotY
    # @option options [String] depthPercent
    # @option options [Boolean] rAngAx
    # @option options [Integer] perspective
    # @see Chart
    # @see View3D
    def initialize(frame, options={})
      super(frame, options)
      @series_type = PieSeries
      @view3D = View3D.new({:rotX=>30, :perspective=>30}.merge(options))
    end

    # Serializes the pie chart
    # @return [String]
    def to_xml
      super() do |xml|
        xml[:c].pie3DChart {
          xml[:c].varyColors :val=>1
          @series.each { |ser| ser.to_xml(xml) }
        }                      
      end
    end
  end
end
