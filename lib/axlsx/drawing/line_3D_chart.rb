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
  #   ws.add_row :values => ["This is a chart with no data in the sheet"]
  #
  #   chart = ws.add_chart(Axlsx::Line3DChart, :start_at=> [0,1], :end_at=>[0,6], :title=>"Most Popular Pets")
  #   chart.add_series :data => [1, 9, 10], :labels => ["Slimy Reptiles", "Fuzzy Bunnies", "Rottweiler"]
  #
  # @see Worksheet#add_chart
  # @see Worksheet#add_row
  # @see Chart#add_series
  # @see Series
  # @see Package#serialize
  class Line3DChart < Chart

    # the category axis
    # @return [CatAxis]
    attr_reader :catAxis

    # the category axis
    # @return [ValAxis]
    attr_reader :valAxis

    # the category axis
    # @return [Axis]
    attr_reader :serAxis

    # space between bar or column clusters, as a percentage of the bar or column width.
    # @return [String]
    attr_reader :gapDepth

    #grouping for a column, line, or area chart.
    # must be one of  [:percentStacked, :clustered, :standard, :stacked]
    # @return [Symbol]
    attr_reader :grouping

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
      @grouping = :standard
      @catAxId = rand(8 ** 8)
      @valAxId = rand(8 ** 8)
      @serAxId = rand(8 ** 8)
      @catAxis = CatAxis.new(@catAxId, @valAxId)
      @valAxis = ValAxis.new(@valAxId, @catAxId)
      @serAxis = SerAxis.new(@serAxId, @valAxId)
      super(frame, options)
      @series_type = LineSeries
      @view3D = View3D.new({:perspective=>30}.merge(options))
    end

    # @see grouping
    def grouping=(v)
      RestrictionValidator.validate "Bar3DChart.grouping", [:percentStacked, :standard, :stacked], v
      @grouping = v
    end

    # @see gapDepth
    def gapDepth=(v)
      RegexValidator.validate "Bar3DChart.gapWidth", GAP_AMOUNT_PERCENT, v
      @gapDepth=(v)
    end

    def to_xml_string(str = '')
      super do |str|
        str << '<line3DChart>'
        str << '<grouping val="' << grouping.to_s << '"/>'
        str << '<varyColors val="1"/>'
        @series.each { |ser| ser.to_xml_str(str) }
        str << '<dLbls>'
        str << '<showLegendKey val="0"/>'
        str << '<showVal val="0"/>'
        str << '<showCatName val="0"/>'
        str << '<showSerName val="0"/>'
        str << '<showPercent val="0"/>'
        str << '<showBubbleSize val="0"/>'
        str << '</dLbls>'
        str << '<gapDepth val="' << @gapDepth.to_s << '"/>' unless @gapDepth.nil?
        str << '<axId val="' << @catAxId.to_s << '"/>'
        str << '<axId val="' << @valAxId.to_s << '"/>'
        str << '<axId val="' << @serAxId.to_s << '"/>'
        str << '</line3DChart>'
        @catAxis.to_xml_str str
        @valAxis.to_xml_str str
        @serAxis.to_xml_str str
      end
    end

    # Serializes the bar chart
    # @return [String]
    def to_xml
      super() do |xml|
        xml.line3DChart {
          xml.grouping :val=>grouping
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
          xml.gapDepth :val=>@gapDepth unless @gapDepth.nil?
          xml.axId :val=>@catAxId
          xml.axId :val=>@valAxId
          xml.axId :val=>@serAxId
        }
        @catAxis.to_xml(xml)
        @valAxis.to_xml(xml)
        @serAxis.to_xml(xml)
      end
    end
  end
end
