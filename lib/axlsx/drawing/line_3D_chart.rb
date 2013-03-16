# encoding: UTF-8
module Axlsx

  # The Line3DChart is a three dimentional line chart (who would have guessed?) that you can add to your worksheet.
  # @example Creating a chart
  #   # This example creates a line in a single sheet.
  #   require "rubygems" # if that is your preferred way to manage gems!
  #   require "axlsx"
  #
  #   p = Axlsx::Package.new
  #   ws = p.workbook.add_worksheet
  #   ws.add_row ["This is a chart with no data in the sheet"]
  #
  #   chart = ws.add_chart(Axlsx::Line3DChart, :start_at=> [0,1], :end_at=>[0,6], :title=>"Most Popular Pets")
  #   chart.add_series :data => [1, 9, 10], :labels => ["Slimy Reptiles", "Fuzzy Bunnies", "Rottweiler"]
  #
  # @see Worksheet#add_chart
  # @see Worksheet#add_row
  # @see Chart#add_series
  # @see Series
  # @see Package#serialize
  class Line3DChart < LineChart

    # space between bar or column clusters, as a percentage of the bar or column width.
    # @return [String]
    attr_reader :gapDepth

    # validation regex for gap amount percent
    GAP_AMOUNT_PERCENT = /0*(([0-9])|([1-9][0-9])|([1-4][0-9][0-9])|500)%/

    # Creates a new line chart object
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
      @gapDepth = nil
      super(frame, options)
      @view_3D = View3D.new({:perspective=>30}.merge(options))
    end

    # @see gapDepth
    def gapDepth=(v)
      RegexValidator.validate "Bar3DChart.gapWidth", GAP_AMOUNT_PERCENT, v
      @gapDepth=(v)
    end

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = '')
      super(str) do |str_inner|
        str_inner << '<c:gapDepth val="' << @gapDepth.to_s << '"/>' unless @gapDepth.nil?
      end
    end
  end
end
