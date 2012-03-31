# encoding: UTF-8
module Axlsx

  # The Bar3DChart is a three dimentional barchart (who would have guessed?) that you can add to your worksheet.
  # @see Worksheet#add_chart
  # @see Chart#add_series
  # @see Package#serialize
  # @see README for an example
  class Bar3DChart < Chart

    # the category axis
    # @return [CatAxis]
    attr_reader :catAxis

    # the valueaxis
    # @return [ValAxis]
    attr_reader :valAxis

    # The direction of the bars in the chart
    # must be one of [:bar, :col]
    # @return [Symbol]
    attr_reader :barDir

    # space between bar or column clusters, as a percentage of the bar or column width.
    # @return [String]
    attr_reader :gapDepth

    # space between bar or column clusters, as a percentage of the bar or column width.
    # @return [String]
    attr_reader :gapWidth

    #grouping for a column, line, or area chart.
    # must be one of  [:percentStacked, :clustered, :standard, :stacked]
    # @return [Symbol]
    attr_reader :grouping

    # The shabe of the bars or columns
    # must be one of  [:cone, :coneToMax, :box, :cylinder, :pyramid, :pyramidToMax]
    # @return [Symbol]
    attr_reader :shape

    # validation regex for gap amount percent
    GAP_AMOUNT_PERCENT = /0*(([0-9])|([1-9][0-9])|([1-4][0-9][0-9])|500)%/

    # Creates a new bar chart object
    # @param [GraphicFrame] frame The workbook that owns this chart.
    # @option options [Cell, String] title
    # @option options [Boolean] show_legend
    # @option options [Symbol] barDir
    # @option options [Symbol] grouping
    # @option options [String] gapWidth
    # @option options [String] gapDepth
    # @option options [Symbol] shape
    # @option options [Integer] rotX
    # @option options [String] hPercent
    # @option options [Integer] rotY
    # @option options [String] depthPercent
    # @option options [Boolean] rAngAx
    # @option options [Integer] perspective
    # @see Chart
    # @see View3D
    def initialize(frame, options={})
      @barDir = :bar
      @grouping = :clustered
      @gapWidth, @gapDepth, @shape = nil, nil, nil
      @catAxId = rand(8 ** 8)
      @valAxId = rand(8 ** 8)
      @catAxis = CatAxis.new(@catAxId, @valAxId)
      @valAxis = ValAxis.new(@valAxId, @catAxId, :tickLblPos => :low)
      super(frame, options)
      @series_type = BarSeries
      @view3D = View3D.new({:rAngAx=>1}.merge(options))
    end

    # The direction of the bars in the chart
    # must be one of [:bar, :col]
    def barDir=(v)
      RestrictionValidator.validate "Bar3DChart.barDir", [:bar, :col], v
      @barDir = v
    end

    #grouping for a column, line, or area chart.
    # must be one of  [:percentStacked, :clustered, :standard, :stacked]
    def grouping=(v)
      RestrictionValidator.validate "Bar3DChart.grouping", [:percentStacked, :clustered, :standard, :stacked], v
      @grouping = v
    end

    # space between bar or column clusters, as a percentage of the bar or column width.
    def gapWidth=(v)
      RegexValidator.validate "Bar3DChart.gapWidth", GAP_AMOUNT_PERCENT, v
      @gapWidth=(v)
    end

    # space between bar or column clusters, as a percentage of the bar or column width.
    def gapDepth=(v)
      RegexValidator.validate "Bar3DChart.gapWidth", GAP_AMOUNT_PERCENT, v
      @gapDepth=(v)
    end

    # The shabe of the bars or columns
    # must be one of  [:cone, :coneToMax, :box, :cylinder, :pyramid, :pyramidToMax]
    def shape=(v)
      RestrictionValidator.validate "Bar3DChart.shape", [:cone, :coneToMax, :box, :cylinder, :pyramid, :pyramidToMax], v
      @shape = v
    end

    def to_xml_string(str = '')
      super do |str|
        str << '<bar3DChart>'
        str << '<barDir val="' << barDir.to_s << '"/>'
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
        str << '<gapWidth val="' << @gapWidth.to_s << '"/>' unless @gapWidth.nil?
        str << '<gapDepth val="' << @gapDepth.to_s << '"/>' unless @gapDepth.nil?
        str << '<shape val="' << @shape.to_s << '"/>'
        str << '<axId val="' << @catAxId.to_s << '"/>'
        str << '<axId val="' << @valAxId.to_s << '"/>'
        str << '<axId val="0"/>'
        str << '</bar3DChart>'
        @catAxis.to_xml_str str
        @valAxis.to_xml_str str
      end
    end
    # Serializes the bar chart
    # @return [String]
    def to_xml
      super() do |xml|
        xml.bar3DChart {
          xml.barDir :val => barDir
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
          xml.gapWidth :val=>@gapWidth unless @gapWidth.nil?
          xml.gapDepth :val=>@gapDepth unless @gapDepth.nil?
          xml.shape :val=>@shape unless @shape.nil?
          xml.axId :val=>@catAxId
          xml.axId :val=>@valAxId
          xml.axId :val=>0
        }
        @catAxis.to_xml(xml)
        @valAxis.to_xml(xml)
      end
    end
  end
end
